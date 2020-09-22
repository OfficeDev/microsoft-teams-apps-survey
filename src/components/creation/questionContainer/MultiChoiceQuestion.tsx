// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { ChoiceContainer, IChoiceContainerOption, IChoiceContainerStrings } from "../../ChoiceContainer";
import "../Creation.scss";
import { Checkbox, Flex, Divider } from "@fluentui/react-northstar";
import { CircleIcon } from "@fluentui/react-icons-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { updateQuestion } from "../../../actions/CreationActions";
import { SurveyUtils } from "../../../utils/SurveyUtils";
import { Localizer } from "../../../utils/Localizer";
import { Constants } from "../../../utils/Constants";

export interface IMCQComponentProps {
    question: actionSDK.ActionDataColumn;
    isValidationModeOn: boolean;
    questionIndex: number;
    onChange?: (props: IMCQComponentProps) => void;
}

export interface IMCQComponentState {
    options: actionSDK.ActionDataColumnOption[];
}
/**
 * MulitChoiceQuestion format, with the option being single select or multiple select
 */
export class MCQComponent extends React.Component<IMCQComponentProps, IMCQComponentState> {

    constructor(props: IMCQComponentProps) {
        super(props);
        this.state = {
            options: JSON.parse(JSON.stringify(this.props.question.options))
        };
    }

    getOptions = () => {
        let choiceOptions = [];
        const choicePrefix = <CircleIcon outline size="small" className="choice-item-circle" />;
        let i = 0;
        this.state.options.forEach((option) => {
            const choiceOption: IChoiceContainerOption = {
                value: option.displayName,
                choicePrefix: choicePrefix,
                choicePlaceholder: Localizer.getString("Choice", (i + 1)),
                deleteChoiceLabel: Localizer.getString("DeleteChoiceX", i + 1)
            };
            choiceOptions.push(choiceOption);
            i++;
        });
        return choiceOptions;
    }

    generateOptionsErrorData(options: actionSDK.ActionDataColumnOption[]): string[] {
        let optionsError: string[] = [];
        if (options.length < 2) { return optionsError; }
        for (let option of options) {
            if (SurveyUtils.isEmptyOrNull(option.displayName)) {
                optionsError.push(Localizer.getString("Required"));
            } else {
                optionsError.push("");
            }
        }
        return optionsError;
    }

    updateQuestion = () => {
        if (JSON.stringify(this.props.question.options) !== JSON.stringify(this.state.options)) {
            let questionCopy: actionSDK.ActionDataColumn = { ...this.props.question };
            questionCopy.options = [...this.state.options];
            updateQuestion(this.props.questionIndex, questionCopy);
        }
    }
/**
 * This function will render the MCQ choice containers with 2 options, user can add more options if needed
*/
    render() {
        let choices = this.getOptions();
        let thisProps: IMCQComponentProps = {
            question: { ...this.props.question },
            isValidationModeOn: this.props.isValidationModeOn,
            questionIndex: this.props.questionIndex
        };
        let optionsError: string[] = ["", ""];
        if (thisProps.isValidationModeOn) {
            optionsError = this.generateOptionsErrorData(thisProps.question.options);
        }
        return (
            <div className="left-space" onBlur={(e) => {
                this.updateQuestion();
            }}>
                <ChoiceContainer
                    optionsError={optionsError}
                    strings={this.getStringsForChoiceComponent()}
                    title=""
                    options={choices}
                    onDeleteChoice={(i) => {
                        let optionsCopy: actionSDK.ActionDataColumnOption[] = [...this.state.options];
                        optionsCopy.splice(i, 1);
                        for (i; i < optionsCopy.length; i++) {
                            optionsCopy[i].name = i.toString();
                        }
                        this.setState(
                            { options: optionsCopy },
                            () => { this.updateQuestion(); }
                        );
                    }}
                    onUpdateChoice={(i, value) => {
                        let optionsCopy: actionSDK.ActionDataColumnOption[] = [...this.state.options];
                        optionsCopy[i].displayName = value;
                        this.setState({
                            options: optionsCopy
                        });
                    }}
                    onAddChoice={() => {
                        let option: actionSDK.ActionDataColumnOption = {
                            name: thisProps.question.options.length.toString(),
                            displayName: ""
                        };
                        let optionsCopy: actionSDK.ActionDataColumnOption[] = [...this.state.options];
                        optionsCopy.push(option);
                        this.setState({
                            options: optionsCopy
                        });
                    }}
                    className="left-zero"
                    limit={10}
                    maxLength={Constants.SURVEY_CHOICE_MAX_LENGTH}
                    inputClassName="invalid-error"
                />
                <Divider className="question-divider" />
                <Flex className="MCQ-setting" gap="gap.large">
                    <Checkbox
                        className="MCQ-setting-item"
                        label={Localizer.getString("MultipleAnswers")}
                        checked={this.props.question.valueType === actionSDK.ActionDataColumnValueType.MultiOption}
                        onChange={(e, data) => {
                            //checkbox to select if the question should be objective or multiple select
                            thisProps.question.valueType = data.checked ? actionSDK.ActionDataColumnValueType.MultiOption : actionSDK.ActionDataColumnValueType.SingleOption;
                            this.props.onChange(thisProps);
                        }} />
                    <Checkbox
                        className="MCQ-setting-item"
                        checked={!this.props.question.allowNullValue}
                        label={Localizer.getString("Required")}
                        onChange={(e, data) => {
                            thisProps.question.allowNullValue = !(data.checked);
                            this.props.onChange(thisProps);
                        }} />
                </Flex>
            </div>
        );
    }

    getStringsForChoiceComponent(): IChoiceContainerStrings {
        let choiceContainerStrings: IChoiceContainerStrings = {
            addChoice: Localizer.getString("AddChoice")
        };
        return choiceContainerStrings;
    }
}
