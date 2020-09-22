// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { IQuestionProps, QuestionView } from "./QuestionView";
import { InputBox,IInputBoxProps } from "./../../InputBox";
import { Localizer } from "../../../utils/Localizer";

export class TextAnswerView extends React.Component<IQuestionProps> {
    render() {
        let value: string = this.props.response as string;
        let props: IInputBoxProps = {
            fluid: true,
            maxLength: 4000,
            multiline: true,
            placeholder: Localizer.getString("EnterAnswer")
        };
        if (value) {
            props.value = value;
        }
        if (this.props.editable) {
            props = {
                ...props,
                defaultValue: value,
                onChange: (e) => {
                    this.props.updateResponse((e.target as HTMLInputElement).value);
                }
            };
        } else {
            props = {
                ...props,
                value: value,
                disabled: true,
                className: "break-word"
            };
        }
        return (
            <QuestionView {...this.props}>
                <InputBox {...props} />
            </QuestionView>
        );
    }
}
