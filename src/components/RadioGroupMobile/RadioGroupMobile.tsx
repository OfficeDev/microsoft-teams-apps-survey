// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";

import "./RadioGroupMobile.scss";
import { CircleIcon, Flex, AcceptIcon } from "@fluentui/react-northstar";

export interface IRadioGroupMobileProps {
    checkedValue: string;
    items: any[];
    checkedValueChanged: (value: string) => void;
}

export class RadioGroupMobile extends React.PureComponent<IRadioGroupMobileProps> {
    render() {
        return (
            <>
                {this.getRadioItemViews()}
            </>
        );
    }

    getRadioItemViews(): JSX.Element[] {
        let radioItemViews: JSX.Element[] = [];
        this.props.items.forEach((item) => {
            let radioItem: JSX.Element = this.getRadioItem(item);
            if (radioItem) {
                radioItemViews.push(radioItem);
            }
        });
        return radioItemViews;
    }

    getRadioItem(item): JSX.Element {
        let isChecked: boolean = item.value == this.props.checkedValue;
        return (
            <Flex className="radio-item-container" key={item.key} onClick={() => { this.props.checkedValueChanged(item.value); }}>
                <div role="radio" aria-checked={isChecked} className={"radio-item-content " + item.className}>{item.label}</div>
                <div className="checkmark-icons-container">
                <CircleIcon className="checkmark-icon checkmark-bg-icon" styles={isChecked ? ({ theme: { siteVariables } }) => ({
                        color: siteVariables.colorScheme.brand.foreground,
                    }) : { color: "grey" }} outline={!isChecked} disabled={!isChecked} />
                    {isChecked ? <AcceptIcon className="checkmark-icon checkmark-tick-icon" style={{ color: "white" }} /> : null}
                </div>
            </Flex>
        );
    }
}
