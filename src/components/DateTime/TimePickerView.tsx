// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Popup, Input, List, ListProps, FocusTrapZone, ChevronDownIcon } from "@fluentui/react-northstar";
import "./TimePickerView.scss";
import { Constants } from "./../../utils/Constants";
import { UxUtils } from "./../../utils/UxUtils";

export interface ITimePickerViewProps {
    placeholder?: string;
    minTimeInMinutes?: number;
    defaultTimeInMinutes?: number;
    renderForMobile?: boolean;
    locale?: string;
    onTimeChange?: (minutes: number) => void;
}

export interface ITimePickerViewState {
    showPicker: boolean;
    selectedTimePickerItem: TimePickerItem;
    timePickerItemsList: TimePickerItem[];
    prevMinTimeInMinutes: number;
}

export class TimePickerView extends React.Component<ITimePickerViewProps, ITimePickerViewState> {
    private timeInputRef: HTMLElement;

    constructor(props: ITimePickerViewProps) {
        super(props);

        let timePickerList: TimePickerItem[] = TimePickerView.getTimePickerList(this.props.minTimeInMinutes, this.props.locale);
        this.state = {
            showPicker: false,
            selectedTimePickerItem: timePickerList.length > 0 ? TimePickerView.getTimePickerListItem(this.props.defaultTimeInMinutes, timePickerList) : null,
            timePickerItemsList: timePickerList,
            prevMinTimeInMinutes: this.props.minTimeInMinutes
        };

        if (this.state.selectedTimePickerItem) {
            this.props.onTimeChange(this.state.selectedTimePickerItem.hours * 60 + this.state.selectedTimePickerItem.minutes);
        }
    }

    static getDerivedStateFromProps(props, state) {
        if (state.prevMinTimeInMinutes == props.minTimeInMinutes) {
            return null;
        }
        let timePickerList: TimePickerItem[] = TimePickerView.getTimePickerList(props.minTimeInMinutes, props.locale);
        return {
            selectedTimePickerItem: TimePickerView.listContainsItem(timePickerList, state.selectedTimePickerItem) ? state.selectedTimePickerItem : timePickerList[0],
            timePickerItemsList: timePickerList,
            prevMinTimeInMinutes: props.minTimeInMinutes
        };
    }

    static getTimePickerList(minTimeInMinutes: number, locale: string): TimePickerItem[] {
        let timePickerList: TimePickerItem[] = [];
        for (let i = 0; i < 24; i++) {
            if (!minTimeInMinutes || i * 60 > minTimeInMinutes) {
                timePickerList.push(new TimePickerItem(i, 0, locale));
            }
            if (!minTimeInMinutes || i * 60 + 30 > minTimeInMinutes) {
                timePickerList.push(new TimePickerItem(i, 30, locale));
            }
        }
        return timePickerList;
    }

    static getTimePickerListItem(timeInMinutes: number, timePickerList: TimePickerItem[]): TimePickerItem {
        let selectedIndex: number = 0;
        for (let i = 0; i < timePickerList.length; i++) {
            let item: TimePickerItem = timePickerList[i];
            if (timeInMinutes <= item.hours * 60 + item.minutes) {
                selectedIndex = i;
                break;
            }
        }
        return timePickerList.length > 0 ? timePickerList[selectedIndex] : null;
    }

    render() {
        return (
            this.props.renderForMobile ?
                this.renderTimePickerForMobile()
                :
                this.renderTimePickerForWebOrDesktop()
        );
    }

    renderTimePickerForMobile() {
        return (
            <>
                {this.renderTimePickerPreviewView()}
                <input
                    ref={(timeInputRef) => {
                        this.timeInputRef = timeInputRef;
                    }}
                    type="time"
                    aria-label={this.state.selectedTimePickerItem.asString}
                    className="hidden-time-input-mob"
                    value={this.state.selectedTimePickerItem.value}
                    onChange={(e) => {
                        let valueInMinutes: number = Math.floor(e.target.valueAsNumber / 60000);
                        if (!this.isTimeValid(valueInMinutes)) {
                            return;
                        }
                        let selectedTime: TimePickerItem = new TimePickerItem(Math.floor(valueInMinutes / 60), valueInMinutes % 60, this.props.locale);
                        this.setState({
                            selectedTimePickerItem: selectedTime
                        });
                        if (this.props.onTimeChange) {
                            this.props.onTimeChange(selectedTime.hours * 60 + selectedTime.minutes);
                        }
                    }}
                    aria-hidden={true}
                />
            </>
        );
    }

    renderTimePickerForWebOrDesktop() {
        let timePickerItems: any[] = [];
        let selectedIndex = this.getSelectedIndex();
        for (let itemIter = 0; itemIter < this.state.timePickerItemsList.length; itemIter++) {
            let item = this.state.timePickerItemsList[itemIter];
            timePickerItems.push({
                key: "tpItem-" + item.asString,
                content: item.asString,
                className: "list-item",
                tabIndex: (itemIter == selectedIndex ? 0 : -1)
            });
        }

        return (
            <Popup
                align="start"
                position="below"
                open={this.state.showPicker}
                onOpenChange={(e, data) => {
                    /*
                    The following isTrusted check is added to prevent any non-user generated events from
                    closing the Popup.
                    When the TimePicker is used within a RadioGroup, like in the case of Notification Settings,
                    when the Enter key is pressed to open the popup, it actually trigers two events: one from
                    the Input element and the other from the underlying radio item. While the first event
                    opens the popup, the second event closes it immediately as it is treated as a click outside the popup.
                    Since the second event is not user generated, the isTrusted flag will be false and we will
                    ignore it here.
                     */
                    if (e.isTrusted) {
                        this.setState({
                            showPicker: data.open
                        });
                    }
                }}
                trigger={this.renderTimePickerPreviewView()}
                content={
                    timePickerItems.length > 0 &&
                    <FocusTrapZone
                        /*
                            This traps the focus within the List component below.
                            On clicking outside the list, the list is dismissed.
                            Special handling is added for Esc key to dismiss the list using keyboard.
                        */
                        onKeyDown={(e) => {
                            if (!e.repeat && (e.keyCode || e.which) == Constants.ESCAPE_ASCII_VALUE && this.state.showPicker) {
                                this.setState({
                                    showPicker: false
                                });
                            }
                        }}>
                        <div className="time-picker-items-list-container" >
                            <List selectable
                                defaultSelectedIndex={selectedIndex}
                                items={timePickerItems}
                                onSelectedIndexChange={(e, props: ListProps) => {
                                    let selectedItem: TimePickerItem = this.state.timePickerItemsList[props.selectedIndex];
                                    this.setState({
                                        showPicker: !this.state.showPicker,
                                        selectedTimePickerItem: selectedItem
                                    });
                                    if (this.props.onTimeChange) {
                                        this.props.onTimeChange(selectedItem.hours * 60 + selectedItem.minutes);
                                    }
                                }}
                            />
                        </div>
                    </FocusTrapZone>
                }
            />
        );
    }

    renderTimePickerPreviewView() {
        let wrapperClassName = "time-input-view time-picker-preview-container";
        if (this.props.renderForMobile) {
            wrapperClassName += " time-input-view-mob";
        }

        let inputWrapperProps = {
            tabIndex: -1,
            "aria-label": (this.props.renderForMobile && this.state.selectedTimePickerItem) ? this.state.selectedTimePickerItem.asString + ". " + this.props.placeholder : null,
            onClick: () => {
                this.onTimePickerPreviewTap();
            },
            className: wrapperClassName,
            "aria-expanded": this.state.showPicker,
            ...UxUtils.getTappableInputWrapperRole()
        };

        let inputProps = {
            type: "text",
            placeholder: this.props.placeholder,
            "aria-hidden": true,
            value: this.state.selectedTimePickerItem.asString,
            readOnly: true,
            "aria-readonly": false,
            className: "time-input"
        };

        return (
            <Input
                input={{ ...inputProps }}
                wrapper={{ ...inputWrapperProps }}
                icon={this.timePickerChevronIcon()}
            />
        );
    }

    timePickerChevronIcon() {
        return <ChevronDownIcon
            outline
            onClick={() => {
                this.onTimePickerPreviewTap();
            }}
        />;
    }

    onTimePickerPreviewTap() {
        if (this.props.renderForMobile && this.timeInputRef) {
            this.timeInputRef.click();
            this.timeInputRef.focus();
        } else {
            this.setState({
                showPicker: !this.state.showPicker
            });
        }
    }

    isTimeValid(minutes: number): boolean {
        if (isNaN(minutes)) {
            return false;
        } else if (this.props.minTimeInMinutes && minutes < this.props.minTimeInMinutes) {
            return false;
        }
        return true;
    }

    getSelectedIndex() {
        let index = 0;
        for (let i = 0; i < this.state.timePickerItemsList.length; i++) {
            if (this.state.timePickerItemsList[i].asString == this.state.selectedTimePickerItem.asString) {
                index = i;
                break;
            }
        }
        return index;
    }

    static listContainsItem(timePickerItemList, item): boolean {
        for (let pickerItem of timePickerItemList) {
            if (pickerItem.value == item.value) {
                return true;
            }
        }
        return false;
    }
}

class TimePickerItem {
    hours: number;
    minutes: number;
    value: string;
    asString: string;
    constructor(hours: number, minutes: number, locale: string = navigator.language) {
        this.hours = hours;
        this.minutes = minutes;

        this.value = this.hours + ":" + this.minutes;
        let date = new Date();
        date.setHours(this.hours);
        date.setMinutes(this.minutes);
        this.asString = date.toLocaleTimeString(locale, { hour: "2-digit", minute: "2-digit", hour12: true });
    }
}
