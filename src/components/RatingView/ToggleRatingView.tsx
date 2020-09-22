// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex } from "@fluentui/react-northstar";
import { LikeIcon } from "@fluentui/react-icons-northstar";
import "./Rating.scss";
import { UxUtils } from "../../utils/UxUtils";
import { Localizer } from "../../utils/Localizer";

export interface IToggleRatingViewProps {
    defaultValue?: boolean;
    disabled?: boolean;
    isPreview?: boolean;
    onChange?: (value: boolean) => void;
}

interface IState {
    value: boolean;
}

export class ToggleRatingView extends React.PureComponent<IToggleRatingViewProps, IState> {

    static getDerivedStateFromProps(props: IToggleRatingViewProps, state) {
        return {
            value: props.defaultValue
        };
    }

    constructor(props: IToggleRatingViewProps) {
        super(props);
        this.state = {
            value: props.defaultValue
        };
    }

    render() {
        let className = "rating-icon";
        if (this.props.disabled) {
            className = className + " disabled-rating";
        } else if (!this.props.isPreview) {
            className = className + " pointer-cursor";
        }
        let isAccessibilityDisabled: boolean = this.props.isPreview || this.props.disabled;
        return (
            <Flex gap="gap.medium">
                <LikeIcon
                    aria-label={this.state.value ? Localizer.getString("LikeTextSelected") : Localizer.getString("LikeText")}
                    {...(!isAccessibilityDisabled) && UxUtils.getTabKeyProps()}
                    outline={this.state.value != true}
                    size="medium"
                    role="button"
                    disabled={this.props.disabled && !this.props.isPreview}
                    aria-disabled={isAccessibilityDisabled}
                    onClick={isAccessibilityDisabled ? null : () => {
                        this.onChange(true);
                    }}
                    className={this.state.value === true ? className : ""} />

                <LikeIcon
                    aria-label={this.state.value === false ? Localizer.getString("DislikeTextSelected") : Localizer.getString("DislikeText")}
                    {...(!isAccessibilityDisabled) && UxUtils.getTabKeyProps()}
                    role="button"
                    outline={this.state.value != false}
                    rotate={180}
                    disabled={this.props.disabled && !this.props.isPreview}
                    aria-disabled={isAccessibilityDisabled}
                    size="medium"
                    onClick={isAccessibilityDisabled ? null : () => {
                        this.onChange(false);
                    }}
                    className={this.state.value === false ? className : ""} />
            </Flex>
        );
    }

    private onChange(value: boolean) {
        if (!this.props.disabled) {
            this.setState({ value: value });
            this.props.onChange(value);
        }
    }
}
