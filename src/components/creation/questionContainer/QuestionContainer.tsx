// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "../Creation.scss";
import { UxUtils } from "../../../utils/UxUtils";
import { Flex } from "@fluentui/react-northstar";
import { CanvasAddPageIcon, TrashCanIcon, ArrowUpIcon, ArrowDownIcon } from "@fluentui/react-icons-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import * as React from "react";
import {
    deleteQuestion,
    duplicateQuestion,
    moveQuestionDown,
    moveQuestionUp,
    updateActiveQuestionIndex,
    updateQuestion,
} from "../../../actions/CreationActions";
import { SurveyUtils } from "../../../utils/SurveyUtils";
import QuestionComponent, { IQuestionComponentProps } from "./QuestionComponent";
import { ResponseViewMode } from "../../../store/ResponseStore";
import ResponseView from "../../Response/ResponseView";
import { Utils } from "../../../utils/Utils";
import { Localizer } from "../../../utils/Localizer";

export interface IQuestionContainerProps {
    questions: actionSDK.ActionDataColumn[];
    activeQuestionIndex: number;
    isValidationModeOn: boolean;
    className?: string;
}

export class QuestionContainer extends React.Component<IQuestionContainerProps> {

    private shouldFocus = false;
    private isQuestionTitleBoxClicked = false;

    shouldComponentUpdate(props: any, nextState: any) {
        //should focus on title only when question title box is clicked and active question index is changed
        //or a new question is added
        if (this.props.activeQuestionIndex !== props.activeQuestionIndex && this.isQuestionTitleBoxClicked
            || this.props.questions.length < props.questions.length) {
            this.shouldFocus = true;
            this.isQuestionTitleBoxClicked = false;
        } else {
            this.shouldFocus = false;
        }
        return true;
    }

    render() {
        const questions: actionSDK.ActionDataColumn[] = this.props.questions;
        let questionsView: JSX.Element[] = [];
        for (let i = 0; i < questions.length; i++) {
            let question: actionSDK.ActionDataColumn = { ...questions[i] };
            if (i === this.props.activeQuestionIndex) {
                questionsView.push(this.getContentView(i, question));
            } else {
                questionsView.push(this.getTitleContentView(i, question));
            }
        }

        return (
            <Flex column>
                {questionsView}
            </Flex>
        );
    }

    private getTitleContentView(index: number, question: actionSDK.ActionDataColumn): JSX.Element {
        let questionPreview: JSX.Element = (
            <div
                key={"question" + index}
                className={(this.props.isValidationModeOn && !SurveyUtils.isQuestionValid(question) ? "questionPaneTitle invalid" : "questionPaneTitle")}
                {...UxUtils.getListItemProps()}
                onClick={(e) => {
                    this.isQuestionTitleBoxClicked = true;
                    updateActiveQuestionIndex(index);
                }}>
                <ResponseView
                    questionNumber={index + 1}
                    actionInstanceColumn={question}
                    responseState={ResponseViewMode.CreationPreview}
                />
            </div>);
        return questionPreview;
    }

    private getContentView(index: number, question: actionSDK.ActionDataColumn) {
        return (
            <div key={"question" + index} className={(this.props.isValidationModeOn && !SurveyUtils.isQuestionValid(question) ? "question-box invalid" : "question-box")}>
                <div className="question-controls">
                    <CanvasAddPageIcon
                        {...UxUtils.getTabKeyProps()}
                        title={Localizer.getString("DuplicateQuestion")}
                        aria-label={Localizer.getString("DuplicateQuestion")}
                        outline xSpacing="after"
                        className="pointer-cursor"
                        onClick={() => {
                            duplicateQuestion(index);
                        }} />

                    <TrashCanIcon
                        {...UxUtils.getTabKeyProps()}
                        title={Localizer.getString("DeleteQuestion")}
                        aria-label={Localizer.getString("DeleteQuestion")}
                        outline
                        xSpacing="after"
                        className="pointer-cursor" onClick={() => {
                            deleteQuestion(index);
                        }} />

                    <ArrowUpIcon
                        {...(index != 0 && UxUtils.getTabKeyProps())}
                        role="button"
                        title={Localizer.getString("MoveQuestionUp")}
                        aria-label={Localizer.getString("MoveQuestionUp")}
                        xSpacing="after"
                        className={index !== 0 ? "pointer-cursor" : ""}
                        disabled={index === 0}
                        aria-disabled={index === 0}
                        onClick={index !== 0 ? () => {
                            moveQuestionUp(index);
                            Utils.announceText("QuestionMovedUp");
                        } : null} />

                    <ArrowDownIcon
                        {...(index != this.props.questions.length - 1 && UxUtils.getTabKeyProps())}
                        role="button"
                        title={Localizer.getString("MoveQuestionDown")}
                        aria-label={Localizer.getString("MoveQuestionDown")}
                        xSpacing="after"
                        className={index !== this.props.questions.length - 1 ? "pointer-cursor" : ""}
                        disabled={index === this.props.questions.length - 1}
                        aria-disabled={index === this.props.questions.length - 1}
                        onClick={index !== this.props.questions.length - 1 ? () => {
                            moveQuestionDown(index);
                            Utils.announceText("QuestionMovedDown");
                        } : null} />
                </div>
                <QuestionComponent
                    isValidationModeOn={this.props.isValidationModeOn}
                    onChange={(props: IQuestionComponentProps) => {
                        updateQuestion(index, props.question);
                    }}
                    question={question}
                    questionIndex={index}
                    shouldFocusOnTitle={index === this.props.activeQuestionIndex && this.shouldFocus}
                    renderForMobile={false}
                />
            </div>
        );
    }
}
