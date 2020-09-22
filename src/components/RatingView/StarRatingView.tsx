// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex } from "@fluentui/react-northstar";
import { StarIcon } from "@fluentui/react-icons-northstar";
import "./Rating.scss";
import { UxUtils } from "../../utils/UxUtils";
import { Utils } from "../../utils/Utils";
import { Localizer } from "../../utils/Localizer";

export interface IStarRatingViewProps {
    max: number;
    icon?: string;
    defaultValue: number;
    disabled?: boolean;
    isPreview?: boolean;
    className?: string;
    onChange?: (value: number) => void;
}

interface IState {
    value: number;
}

export class StarRatingView extends React.PureComponent<IStarRatingViewProps, IState> {

    constructor(props: IStarRatingViewProps) {
        super(props);
        this.state = {
            value: props.defaultValue
        };
    }

    static getDerivedStateFromProps(props: IStarRatingViewProps, state) {
        return {
            value: props.defaultValue
        };
    }

    render() {
        let items: JSX.Element[] = [];
        for (let i = 1; i <= this.props.max; i++) {
            let className = this.state.value < i ? "rating-icon-unfilled" : "rating-icon";
            className = (this.props.disabled && this.state.value >= i) ? className + " disabled-rating" : className;
            if (!this.props.isPreview && !this.props.disabled) {
                className = className + " pointer-cursor";
            }
            let isAccessibilityDisabled: boolean = this.props.disabled || this.props.isPreview;
            items.push(
                <StarIcon
                    {...(!isAccessibilityDisabled) && UxUtils.getTabKeyProps()}
                    aria-label={i <= this.state.value ? Localizer.getString("StarValueSelected", i) : Localizer.getString("StarValue", i)}
                    key={i}
                    outline={this.state.value < i}
                    disabled={this.props.disabled && !this.props.isPreview}
                    aria-disabled={isAccessibilityDisabled}
                    onClick={isAccessibilityDisabled ? null : () => {
                        Utils.announceText(Localizer.getString("StarNumberSelected", i));
                        this.setState({ value: i });
                        this.props.onChange(i);
                    }}
                    className={className}
                />
            );
        }
        return (
            <Flex gap="gap.medium" className={this.props.className}>
                {items}
            </Flex>
        );
    }
}
