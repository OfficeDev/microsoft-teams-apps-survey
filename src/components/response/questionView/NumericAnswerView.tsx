// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { IQuestionProps, QuestionView } from "./QuestionView";
import { Input, InputProps } from "@fluentui/react-northstar";
import { Localizer } from "../../../utils/Localizer";

export class NumericAnswerView extends React.Component<IQuestionProps> {

    //The max value for numeric is taken based on the most significant digit allowed for INT64 value
    readonly SURVEY_NUMERIC_MAX_VALUE = 999999999999999;

    shouldComponentUpdate(nextProps) {
        if (isNaN(nextProps.response)) {
            return false;
        }
        return true;
    }

    render() {
        let props: InputProps = {
            placeholder: Localizer.getString("EnterNumber"),
            type: "number",
            fluid: true,
            // The max and maxLength of the numeric answer is set based on the most significant bit an INT64 value can take
            maxLength: 15,
            max: this.SURVEY_NUMERIC_MAX_VALUE,
            required: this.props.required //adding required field to be able to capture valueMissing error
        };
        let value: string = isNaN(parseFloat(this.props.response)) ? "" : this.props.response as string;
        if (value) {
            props.value = value;
        }
        if (this.props.editable) {
            props = {
                ...props,
                defaultValue: value,
                onChange: (event) => {
                    //Onchange event is needed when only single response is allowed and user wants to edit the previous filled survey
                   if ((event.currentTarget as HTMLInputElement).validity.badInput) {
                        this.props.updateResponse("badInput");
                    } else if ((event.currentTarget as HTMLInputElement).validity.valueMissing) {
                        //saving "" in store whenever currentTarget validity.valueMissing is true
                        //so that empty response validation is taken care of
                        this.props.updateResponse("");
                    } else if((Number((event.currentTarget as HTMLInputElement).value)) > this.SURVEY_NUMERIC_MAX_VALUE) {
                        this.props.updateResponse("badInput");
                    } else {
                        this.props.updateResponse(Number((event.currentTarget as HTMLInputElement).value).toString());
                    }
                },
                onBlur: (event) => {
                    if((Number((event.currentTarget as HTMLInputElement).value)) > this.SURVEY_NUMERIC_MAX_VALUE) {
                         this.props.updateResponse("badInput");
                    } else {
                            this.props.updateResponse(Number((event.currentTarget as HTMLInputElement).value).toString());
                    }
                 }
            };
        } else {
            props = {
                ...props,
                value: value,
                disabled: true
            };
        }
        return (
            <QuestionView {...this.props}>
                <Input {...props} />
            </QuestionView>
        );
    }
}
