// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { QuestionView, IMultiChoiceProps } from "./QuestionView";
import { RadioGroup } from "@fluentui/react-northstar";
import { CircleIcon } from "@fluentui/react-icons-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import "../Response.scss";

export class SingleSelectView extends React.Component<IMultiChoiceProps> {

    render() {
        let response: string = this.props.response as string;
        return (
            <QuestionView {...this.props}>
                <RadioGroup
                    vertical
                    items={this.getItems(this.props.options)}
                    checkedValue={!this.props.isPreview ? response : -1}
                    onCheckedValueChange={!this.props.isPreview ? (event, data) => {
                        this.props.updateResponse(data.value);
                    } : null} />
            </QuestionView>
        );
    }

    private getItems(options: actionSDK.ActionDataColumnOption[]): any[] {
        let opts: any[] = [];
        let className = "single-select";
        if (this.props.editable) {
            className = className + " pointer-cursor-important";
        }
        for (let i = 0; i < options.length; i++) {
            opts.push({
                key: options[i].name,
                label: options[i].displayName,
                value: options[i].name,
                icon: <CircleIcon
                    size="medium"
                    aria-disabled={!this.props.editable}
                    className={(this.props.response == i && !this.props.editable) ? "icon-disabled" : ""} />,
                disabled: !this.props.editable && !this.props.isPreview,
                "aria-disabled": !this.props.editable,
                className: i !== options.length - 1 ? "single-select options-space" : "single-select",
                tabIndex: this.props.editable ? 0 : -1,
                role: "radio"
            });

        }

        return opts;
    }
}
