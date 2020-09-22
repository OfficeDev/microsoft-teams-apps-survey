// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Text, Flex } from "@fluentui/react-northstar";
import "../Response.scss";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { UxUtils } from "./../../../utils/UxUtils";
import { Localizer } from "../../../utils/Localizer";

export interface IRatingAnswerProps extends IQuestionProps {
    count: number;
}

export interface IMultiChoiceProps extends IQuestionProps {
    options: actionSDK.ActionDataColumnOption[];
}

export interface IQuestionProps {
    questionNumber: number;
    questionText: string;
    required?: boolean;
    editable?: boolean;
    response?: any;
    isPreview?: boolean;
    locale?: string;
    updateResponse?: (response: any) => void;
}

export class QuestionView extends React.Component<IQuestionProps> {

    render() {
        let className = "question-view-title break-word";
        return (
            <Flex gap="gap.small" {...this.getAccessibilityProps()}>
                <Flex className="question-number-text">
                    <Text content={this.props.questionNumber + ". "} className="question-view-title" />
                </Flex>
                <Flex gap="gap.smaller" className="question-view-content" column fill>
                    {this.props.required ?
                        <div aria-label={this.getQuestionText() + " " + Localizer.getString("Required")}>
                            <Text className={className} content={this.getQuestionText()} aria-hidden={true} />
                            <span className="required-color" aria-hidden={true}> *</span>
                        </div>
                        : <Text className={className} content={this.getQuestionText()} />}
                    {this.props.children}
                </Flex>
            </Flex>
        );
    }

    private getQuestionText = () => {
        if (this.props.questionText) {
            return this.props.questionText;
        }
        return Localizer.getString("QuestionTitlePlaceHolder");
    }

    //Adding this prop to stop announcing 2 times on android phone
    private getAccessibilityProps = () => {
        if (UxUtils.renderingForAndroid()) {
            return {
                role: "group"
            };
        }
    }
}
