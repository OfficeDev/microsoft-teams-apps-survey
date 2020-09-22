// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Checkbox, Flex, Input, Text, Divider } from "@fluentui/react-northstar";
import { CalendarIcon } from "@fluentui/react-icons-northstar";
import { MCQComponent } from "./MultiChoiceQuestion";
import { RatingsQuestionComponent } from "./RatingsQuestionComponent";
import { QuestionDisplayType } from "./QuestionDisplayType";
import * as actionSDK from "@microsoft/m365-action-sdk";
import "../Creation.scss";
import { InputBox } from "../../InputBox";
import { updateQuestion } from "../../../actions/CreationActions";
import { observer } from "mobx-react";
import { Localizer } from "../../../utils/Localizer";
import { Constants } from "../../../utils/Constants";

export interface IQuestionComponentProps {
    question: actionSDK.ActionDataColumn;
    isValidationModeOn: boolean;
    questionIndex: number;
    shouldFocusOnTitle?: boolean;
    renderForMobile: boolean;
    onChange?: (props: IQuestionComponentProps) => void;
}
/**
* Question component which defines/gets the question formatting based in the type user wants to add
*/
@observer
export default class QuestionComponent extends React.Component<IQuestionComponentProps> {
    constructor(props: IQuestionComponentProps) {
        super(props);
    }

    getDisplayType(question: actionSDK.ActionDataColumn) {
        let customProperties = JSON.parse(question.properties);
        if (customProperties && customProperties.hasOwnProperty("dt")) {
            return customProperties.dt;
        }
        return QuestionDisplayType.None;
    }
    /**
    * Get the question view for all the question types
    */
    getQuestionView() {
        let thisProps: IQuestionComponentProps = {
            question: { ...this.props.question },
            isValidationModeOn: this.props.isValidationModeOn,
            questionIndex: this.props.questionIndex,
            renderForMobile: this.props.renderForMobile
        };
        if (this.props.question.valueType === actionSDK.ActionDataColumnValueType.SingleOption) {
            let displayType = this.getDisplayType(this.props.question);
            if (displayType === QuestionDisplayType.FiveStar ||
                displayType === QuestionDisplayType.TenStar ||
                displayType === QuestionDisplayType.LikeDislike ||
                displayType === QuestionDisplayType.FiveNumber ||
                displayType === QuestionDisplayType.TenNumber) {
                return (
                    <RatingsQuestionComponent renderForMobile={thisProps.renderForMobile}
                        onChange={(props: IQuestionComponentProps) => {
                            this.props.onChange(props);
                        }} question={this.props.question} questionIndex={thisProps.questionIndex} />
                );
            }
            if (displayType === QuestionDisplayType.Select) { /*MCQ question to allow only single option selection */
                return (
                    <MCQComponent isValidationModeOn={this.props.isValidationModeOn} onChange={(props: IQuestionComponentProps) => {
                        this.props.onChange(props);
                    }} question={this.props.question} questionIndex={this.props.questionIndex}>
                    </MCQComponent>
                );
            }
        }
        if (this.props.question.valueType === actionSDK.ActionDataColumnValueType.MultiOption) {  /*MCQ question to allow multiple option selection */
            return (
                <MCQComponent isValidationModeOn={this.props.isValidationModeOn} onChange={(props: IQuestionComponentProps) => {
                    this.props.onChange(props);
                }} question={this.props.question} questionIndex={this.props.questionIndex}>
                </MCQComponent>
            );
        }

        if (this.props.question.valueType === actionSDK.ActionDataColumnValueType.Numeric) {
            return this.getQuestionBase(Localizer.getString("EnterNumber"), thisProps);
        }
        if (this.props.question.valueType === actionSDK.ActionDataColumnValueType.Date) {
            return this.getQuestionBase(Localizer.getString("EnterDate"), thisProps, <CalendarIcon outline={true} />);
        }
        if (this.props.question.valueType === actionSDK.ActionDataColumnValueType.Text) {
            return this.getQuestionBase(Localizer.getString("EnterAnswer"), thisProps);
        }

        return (<Checkbox checked={!(this.props.question.allowNullValue)} label={Localizer.getString("Required")} onChange={(e, data) => {
            thisProps.question.allowNullValue = !(data.checked);
            this.props.onChange(thisProps);
        }} />);
    }

    render() {
        let question: actionSDK.ActionDataColumn = { ...this.props.question };
        return (
            <Flex gap="gap.smaller" className="question-component">
                {!this.props.renderForMobile && <Text content={(this.props.questionIndex + 1) + "."} weight="bold" />}
                <Flex column fill className="zero-min-width">
                    <div className="question-text">
                        <InputBox
                            ref={(inputBox) => {
                                if (inputBox && this.props.shouldFocusOnTitle) {
                                    setTimeout(() => {
                                        inputBox.focus();
                                    }, 0);
                                }
                            }}
                            maxLength={Constants.SURVEY_QUESTION_MAX_LENGTH}
                            className={(this.props.isValidationModeOn && question.displayName.length == 0 ? "invalid" : "")}
                            fluid
                            key={this.props.questionIndex + question.displayName}
                            defaultValue={question.displayName}
                            placeholder={Localizer.getString("QuestionTitlePlaceHolder")}
                            onBlur={(e) => {
                                if ((e.target as HTMLInputElement).value !== question.displayName) {
                                    question.displayName = (e.target as HTMLInputElement).value;
                                    updateQuestion(this.props.questionIndex, question);
                                }
                            }
                            }
                            showError={(this.props.isValidationModeOn && question.displayName.length == 0)}
                            errorText={Localizer.getString("EmptyQuestionTitle")}
                            input={{
                                className: (this.props.isValidationModeOn && question.displayName.length == 0 ? "invalid-error" : "")
                            }}
                        />
                    </div>
                    {this.getQuestionView()}
                </Flex>
            </Flex>
        );
    }
    /**
    * Question component used for question type: text/number/date
    */
    private getQuestionBase(placeholder: string, thisProps: any, icon?: any) {
        return (
            <Flex column className="question-base" gap="gap.medium">
                <Input disabled placeholder={placeholder} fluid icon={icon} className="question-item" />
                <Divider className="question-divider" />
                <Checkbox checked={!(this.props.question.allowNullValue)} label={Localizer.getString("Required")} onChange={(e, data) => {
                    thisProps.question.allowNullValue = !(data.checked);
                    this.props.onChange(thisProps);
                }} className="required-question-checkbox" />
            </Flex>
        );
    }
}
