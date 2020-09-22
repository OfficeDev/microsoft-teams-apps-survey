// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex, Text, ChevronStartIcon, Checkbox,RadioGroup } from "@fluentui/react-northstar";
import { UxUtils } from "../../utils/UxUtils";
import { Localizer } from "../../utils/Localizer";
import { DateTimePickerView } from "../DateTime";
import * as actionSDK from "@microsoft/m365-action-sdk";
import "./Settings.scss";

/**
* These are the two custom parameters provided in survey app, which can be set/modiifed by creator of survey
*/
export enum SettingsSections {
    DUE_BY,
    RESULTS_VISIBILITY,
    MULTI_RESPONSE
}

export interface ISettingsComponentProps {
    dueDate: number;
    locale?: string;
    resultVisibility: actionSDK.Visibility;
    isResponseEditable: boolean;
    isResponseAnonymous: boolean;
    renderForMobile?: boolean;
    isMultiResponseAllowed?: boolean;
    strings: ISettingsComponentStrings;
    renderDueBySection?: () => React.ReactElement<any>;
    renderResultVisibilitySection?: () => React.ReactElement<any>;
    renderNotificationsSection?: () => React.ReactElement<any>;
    renderResponseOptionsSection?: () => React.ReactElement<any>;
    onChange?: (props: ISettingsComponentProps) => void;
    onMount?: () => void;
    onBack?: () => void;
}

export interface ISettingsComponentStrings {
    dueBy?: string;
    multipleResponses?: string;
    responseOptions?: string;
    resultsVisibleTo?: string;
    resultsVisibleToAll?: string;
    resultsVisibleToSender?: string;
    datePickerPlaceholder?: string;
    timePickerPlaceholder?: string;
}

export class Settings extends React.PureComponent<ISettingsComponentProps> {

    private settingProps: ISettingsComponentProps;
    constructor(props: ISettingsComponentProps) {
        super(props);
    }

    componentDidMount() {
        if (this.props.onMount) {
            this.props.onMount();
        }
    }

    render() {
        this.settingProps = {
            dueDate: this.props.dueDate,
            locale: this.props.locale,
            resultVisibility: this.props.resultVisibility,
            isResponseAnonymous: this.props.isResponseAnonymous,
            isResponseEditable: this.props.isResponseEditable,
            isMultiResponseAllowed: this.props.hasOwnProperty("isMultiResponseAllowed") ? this.props.isMultiResponseAllowed : false,
            strings: this.props.strings
        };
        if (this.props.renderForMobile) {
            return this.renderSettings();
        } else {
            return (
                <Flex className="body-container" column gap="gap.medium">
                    {this.renderSettings()}
                    {this.props.onBack && this.getBackElement()}
                </Flex>
            );
        }

    }

    renderSettings() {
        return (
            <Flex column>
                {this.renderDueBySection()}
                {this.renderResultVisibilitySection()}
                {this.renderResponseOptionsSection()}
            </Flex>
        );
    }

    renderDueBySection() {
        let dueByClassName = this.props.renderForMobile ? "due-by-pickers-container date-time-equal" : "settings-indentation" ;
        if (this.props.renderDueBySection) {
            return this.props.renderDueBySection();
        } else {
            return (
                <Flex className="settings-item-margin" role="group" aria-label={this.getString("dueBy")} column gap="gap.smaller">
                    <label className="settings-item-title">{this.getString("dueBy")}</label>
                    <div className={dueByClassName}>
                        <DateTimePickerView showTimePicker
                            minDate={new Date()}
                            locale={this.props.locale}
                            value={new Date(this.props.dueDate)}
                            placeholderDate={this.getString("datePickerPlaceholder")}
                            placeholderTime={this.getString("timePickerPlaceholder")}
                            renderForMobile={this.props.renderForMobile}
                            onSelect={(date: Date) => {
                                this.settingProps.dueDate = date.getTime();
                                this.props.onChange(this.settingProps);
                            }} />
                    </div>
                </Flex>
            );
        }
    }

    renderResultVisibilitySection() {
        if (this.props.renderResultVisibilitySection) {
            return this.props.renderResultVisibilitySection();
        } else {
            return (
                <Flex
                    className="settings-item-margin"
                    role="group"
                    aria-label={this.getString("resultsVisibleTo")}
                    column gap="gap.smaller">
                    <label className="settings-item-title">{this.getString("resultsVisibleTo")}</label>
                    <div className="settings-indentation">
                        <RadioGroup
                            vertical
                            checkedValue={this.settingProps.resultVisibility}
                            items={Settings.getVisibilityItems(this.getString("resultsVisibleToAll"), this.getString("resultsVisibleToSender"))}
                            onCheckedValueChange={(e, props) => {
                                this.settingProps.resultVisibility = props.value as actionSDK.Visibility;
                                this.props.onChange(this.settingProps);
                            }}
                        />
                    </div>
                </Flex>
            );
        }
    }

    renderResponseOptionsSection() {
        let multiOptionClassName = "settings-indentation multi-response";
        if (this.props.renderResponseOptionsSection) {
            return this.props.renderResponseOptionsSection();
        } else {
            return (
                <Flex className="settings-item-margin" role="group" aria-label={this.getString("responseOptions")} column gap="gap.small">
                    <label className="settings-item-title">{this.getString("responseOptions")}</label>
                    <Checkbox
                        role="checkbox"
                        className={multiOptionClassName}
                        checked={this.props.isMultiResponseAllowed}
                        label={this.getString("multipleResponses")}
                        onChange={(e, props) => {
                            this.settingProps.isMultiResponseAllowed = props.checked;
                            this.props.onChange(this.settingProps);
                        }} />
                </Flex>
            );
        }
    }

    getString(key: string): string {
        if (this.props.strings && this.props.strings.hasOwnProperty(key)) {
            return this.props.strings[key];
        }
        return key;
    }

    private getBackElement() {
        if (true /*!this.props.renderForMobile*/) {
            return (
                <Flex className="footer-layout" gap={"gap.smaller"}>
                    <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()} onClick={() => {
                        this.props.onBack();
                    }} >
                        <ChevronStartIcon xSpacing="after" size="small" />
                        <Text content={Localizer.getString("Back")} />
                    </Flex>
                </Flex>
            );
        }
    }

    public static shouldRenderSection(section: SettingsSections, excludedSections: SettingsSections[]) {
        return !excludedSections || (excludedSections.indexOf(section) == -1);
    }

    public static getVisibilityItems(resultsVisibleToAllLabel: string, resultsVisibleToSenderLabel: string) {
        return [
            {
                key: "1",
                label: resultsVisibleToAllLabel,
                value: actionSDK.Visibility.All,
                className: "settings-radio-item"
            },
            {
                key: "2",
                label: resultsVisibleToSenderLabel,
                value: actionSDK.Visibility.Sender,
                className: "settings-radio-item-last"
            },
        ];
    }
}
