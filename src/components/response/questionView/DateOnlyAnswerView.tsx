// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { QuestionView, IQuestionProps } from "./QuestionView";
import { DateTimePickerView, IDateTimePickerViewProps } from "./../../DateTime";
import { Localizer } from "../../../utils/Localizer";
import { UxUtils } from "./../../../utils/UxUtils";

export class DateOnlyAnswerView extends React.Component<IQuestionProps> {

    render() {

        let dateProps: IDateTimePickerViewProps = {
            placeholderDate: Localizer.getString("SelectADate"),
            disabled: !this.props.editable,
            renderForMobile: UxUtils.renderingForMobile(),
            onSelect: (date: Date) => {
                if (date) {
                    this.props.updateResponse(date.toLocaleDateString("en-US"));
                }
            },
            isPreview: this.props.isPreview,
            locale: this.props.locale
        };
        if (this.props.response) {
            dateProps.value = new Date(this.props.response);
        }
        return (
            <QuestionView {...this.props}>
                <DateTimePickerView {...dateProps} />
            </QuestionView>
        );
    }
}
