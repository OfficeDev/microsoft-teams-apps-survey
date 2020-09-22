// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "./ChoiceContainer.scss";
import { InputBox } from "../InputBox";
import { Text, ShorthandValue, AddIcon, BoxProps, TrashCanIcon } from "@fluentui/react-northstar";
import { UxUtils } from "./../../utils/UxUtils";
import { Constants } from "./../../utils/Constants";

export interface IChoiceContainerOption {
    value: string;
    choicePrefix?: JSX.Element;
    choicePlaceholder: string;
    deleteChoiceLabel: string;
}

export interface IChoiceContainerStrings {
    addChoice: string;
}

export interface IChoiceContainerProps {
    title?: string;
    options: IChoiceContainerOption[];
    optionsError?: string[];
    limit?: number;
    renderForMobile?: boolean;
    focusOnError?: boolean;
    strings: IChoiceContainerStrings;
    inputClassName?: string;
    onUpdateChoice?: (i, value) => void;
    onDeleteChoice?: (i) => void;
    onAddChoice?: () => void;
    className?: string;
    maxLength?: number;
}

export class ChoiceContainer extends React.PureComponent<IChoiceContainerProps> {

    private currentFocus: number = -1;
    private addButtonRef: HTMLElement;

    constructor(props: IChoiceContainerProps) {
        super(props);
    }

    getDeleteIconProps(i: number): ShorthandValue<BoxProps> {
        if (this.props.options.length > 2) {
            return {
                content: <TrashCanIcon className="choice-trash-can" outline={true} aria-hidden="false" title={this.props.options[i].deleteChoiceLabel}
                    onClick={() => {
                        if (this.currentFocus == this.props.options.length - 1) {
                            setTimeout((() => {
                                this.addButtonRef.focus();
                            }).bind(this), 0);
                        }
                        this.props.onDeleteChoice(i);
                    }} />,
                ...UxUtils.getTabKeyProps()
            };
        }
        return null;
    }

    render() {
        let items: JSX.Element[] = [];

        let maxOptions: number = (this.props.limit && this.props.limit > 0) ? this.props.limit : Number.MAX_VALUE;
        let focusOnErrorSet: boolean = false;
        let className: string = ("item-content" + ((this.props.options.length > 2) ? " icon-padding" : ""));
        for (let i = 0; i < (maxOptions > this.props.options.length ? this.props.options.length : maxOptions); i++) {
            let errorString = this.props.optionsError && this.props.optionsError.length > i ? this.props.optionsError[i] : "";
            if (errorString.length > 0 && this.props.focusOnError && !focusOnErrorSet) {
                this.currentFocus = i;
                focusOnErrorSet = true;
            }
            if (errorString.length > 0 && this.props.inputClassName) {
                className = className + " " + this.props.inputClassName;
            }
            items.push(
                <div key={"option" + i} className="choice-item">
                    <InputBox
                        ref={(inputBox) => {
                            if (inputBox && i == this.currentFocus) {
                                inputBox.focus();
                            }
                        }}
                        fluid
                        input={{ className }}
                        maxLength={this.props.maxLength}
                        icon={this.getDeleteIconProps(i)}
                        showError={errorString.length > 0}
                        errorText={errorString}
                        key={"option" + i}
                        value={this.props.options[i].value}
                        placeholder={this.props.options[i].choicePlaceholder}
                        onKeyDown={(e) => {
                            if (!e.repeat && (e.keyCode || e.which) == Constants.CARRIAGE_RETURN_ASCII_VALUE && this.props.options.length < maxOptions) {
                                if (i == this.props.options.length - 1) {
                                    this.props.onAddChoice();
                                    this.currentFocus = this.props.options.length;
                                } else {
                                    this.currentFocus += 1;
                                    this.forceUpdate();
                                }
                            }
                        }}
                        onFocus={(e) => {
                            this.currentFocus = i;
                        }}
                        onChange={(e) => {
                            this.props.onUpdateChoice(i, (e.target as HTMLInputElement).value);
                        }}
                        prefixJSX={this.props.options[i].choicePrefix}
                    />
                </div>
            );
        }
        return (
            <div
                className="choice-container"
                onBlur={(e) => {
                    this.currentFocus = -1;
                }}>
                {this.props.title && <div className={this.getChoiceTitleClassName()}>{this.props.title}</div>}
                {items}
                {this.props.options.length < maxOptions &&
                    <div ref={(e) => {
                        this.addButtonRef = e;
                    }} className={this.props.className ? this.props.className + " add-options" : "add-options"} {...UxUtils.getTabKeyProps()} onClick={(e) => {
                        this.props.onAddChoice();
                        this.currentFocus = this.props.options.length;
                    }}>
                        <AddIcon className="plus-icon" outline size="medium" color="brand" />
                        <Text size="medium" content={this.props.strings.addChoice ? this.props.strings.addChoice : "Add Choice"} color="brand" />
                    </div>
                }
            </div>
        );
    }

    getChoiceTitleClassName() {
        return this.props.renderForMobile ? "choice-title-mob" : "choice-title";
    }
}
