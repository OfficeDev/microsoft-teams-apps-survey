// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";

import * as actionSDK from "@microsoft/m365-action-sdk";
import { INavBarComponentProps, NavBarItemType, NavBarComponent } from "../NavBarComponent";
import { goToPage, updateQuestion, deleteQuestion } from "../../actions/CreationActions";
import getStore, { Page } from "../../store/CreationStore";
import { Flex, Text, Divider, Button } from "@fluentui/react-northstar";
import { TrashCanIcon } from "@fluentui/react-icons-northstar";
import { observer } from "mobx-react";
import { SurveyUtils } from "../../utils/SurveyUtils";
import QuestionComponent, { IQuestionComponentProps } from "./questionContainer/QuestionComponent";
import { toJS } from "mobx";
import { Localizer } from "../../utils/Localizer";

@observer
export class UpdateQuestionPage extends React.Component<any> {
    private questionIndex: number;
    private currentActiveIndex: number = -1;
    private shouldFocusOnTitle: boolean = false;

    render() {
        this.questionIndex = getStore().activeQuestionIndex;

        if (this.questionIndex !== this.currentActiveIndex) {
            this.shouldFocusOnTitle = true;
            this.currentActiveIndex = getStore().activeQuestionIndex;
        } else {
            this.shouldFocusOnTitle = false;
        }

        return (
            <Flex className="body-container no-mobile-footer" column>
                {this.getNavBar()}
                {this.getQuestionSection()}
                {this.getDeleteQuestionButton()}
            </Flex>
        );
    }

    private getNavBar() {
        let navBarComponentProps: INavBarComponentProps = {
            title: Localizer.getString("QuestionIndex", this.questionIndex + 1),
            rightNavBarItem: {
                title: Localizer.getString("Done").toUpperCase(),
                ariaLabel: Localizer.getString("Done"),
                onClick: () => {
                    // React Bug: Tapping outside a React component's hierarchy in React doesn't invoke onBlur on the active element
                    // https://github.com/moroshko/react-autosuggest/issues/380
                    // We are dependent on onBlur events to update our stores for input elements.
                    // e.g: Tapping on Send Survey/Create Poll buttons in the Nav Bar doesn't invoke onBlur on a focused input element
                    (document.activeElement as HTMLElement).blur();
                    goToPage(Page.Main);
                },
                type: NavBarItemType.BACK,
                className: "nav-bar-done"
            }
        };
        return (
            <NavBarComponent {...navBarComponentProps} />
        );
    }

    private getQuestionSection() {
        let question: actionSDK.ActionDataColumn = toJS(getStore().questions[this.questionIndex]);

        return (
            <div key={"question" + this.questionIndex} className={(getStore().isValidationModeOn && !SurveyUtils.isQuestionValid(question) ? "question-box invalid" : "question-box")}>
                <QuestionComponent
                    isValidationModeOn={getStore().isValidationModeOn}
                    onChange={(props: IQuestionComponentProps) => {
                        updateQuestion(this.questionIndex, props.question);
                    }}
                    question={question}
                    questionIndex={this.questionIndex}
                    shouldFocusOnTitle={this.shouldFocusOnTitle}
                    renderForMobile={true}
                />
            </div>
        );
    }

    private getDeleteQuestionButton() {
        return (
            <>
                <Divider className="delete-button-divider" />
                <Button
                    className="delete-question-button"
                    text
                    content={
                        <Flex vAlign="center" className="delete-question-button-container">
                            <TrashCanIcon outline className="delete-button-icon" />
                            <Text content={Localizer.getString("DeleteQuestion")} className="delete-button-label" />
                        </Flex>
                    }

                    onClick={() => {
                        deleteQuestion(this.questionIndex);
                    }}
                />
            </>
        );
    }
}
