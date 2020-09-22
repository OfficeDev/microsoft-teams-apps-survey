// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex } from "@fluentui/react-northstar";
import { DatePickerView, IDatePickerViewProps } from "./DatePickerView";
import { TimePickerView, ITimePickerViewProps } from "./TimePickerView";

export interface IDateTimePickerViewProps {
    placeholderDate?: string;
    placeholderTime?: string;
    value?: Date;
    minDate?: Date;
    disabled?: boolean;
    showTimePicker?: boolean;
    renderForMobile?: boolean;
    isPreview?: boolean;
    locale?: string;
    onSelect?: (date: Date) => void;
}

export interface IDateTimePickerViewState {
    selectedDate: Date;
    selectedTime: number;
}

export class DateTimePickerView extends React.Component<IDateTimePickerViewProps, IDateTimePickerViewState> {

    constructor(props: IDateTimePickerViewProps) {
        super(props);

        this.state = {
            selectedDate: this.props.value,
            selectedTime: DateTimePickerView.getTimeInMinutes(this.props.value ? this.props.value : new Date())
        };

        this.props.onSelect(this.state.selectedDate);
    }

    static getDerivedStateFromProps(props, state) {
        return {
            selectedDate: props.value,
            selectedTime: DateTimePickerView.getTimeInMinutes(props.value ? props.value : new Date())
        };
    }

    render() {
        let props: IDatePickerViewProps = {
            placeholder: this.props.placeholderDate,
            date: this.state.selectedDate,
            minDate: this.props.minDate,
            disabled: this.props.disabled,
            locale: this.props.locale,
            renderForMobile: this.props.renderForMobile,
            onSelectDate: (newDate: Date) => {
                if (!this.props.isPreview) {
                    this.dateSelectCallback(newDate);
                }
            }
        };

        let timePickerProps: ITimePickerViewProps = {
            placeholder: this.props.placeholderTime,
            minTimeInMinutes: this.getMinTimeInMinutes(this.state.selectedDate),
            defaultTimeInMinutes: DateTimePickerView.getTimeInMinutes(this.state.selectedDate),
            renderForMobile: this.props.renderForMobile,
            onTimeChange: (minutes: number) => {
                this.timeSelectCallback(minutes);
            },
            locale: this.props.locale
        };

        return (
            <Flex gap={this.props.renderForMobile ? null : "gap.small"} space={this.props.renderForMobile ? "between" : null}>
                <DatePickerView {...props} />
                {this.props.showTimePicker ? <TimePickerView {...timePickerProps} /> : null}
            </Flex>
        );
    }

    dateSelectCallback(newDate: Date) {
        let updatedDate = newDate;
        if (this.props.showTimePicker) {
            if (this.getMinTimeInMinutes(newDate) <= this.state.selectedTime) {
                updatedDate.setHours(Math.floor(this.state.selectedTime / 60));
                updatedDate.setMinutes(this.state.selectedTime % 60);
                this.setState({
                    selectedDate: updatedDate
                });
            } else {
                let updatedHours = Math.floor(this.getMinTimeInMinutes(newDate) / 60);
                let updatedMinutes = this.getMinTimeInMinutes(newDate) % 60;
                if (updatedMinutes > 0 && updatedMinutes <= 30) {
                    updatedMinutes = 30;
                } else if (updatedMinutes > 31) {
                    updatedHours += 1;
                    updatedMinutes = 0;
                }
                updatedDate.setHours(updatedHours);
                updatedDate.setMinutes(updatedMinutes);
                this.setState({
                    selectedDate: updatedDate,
                    selectedTime: updatedHours * 60 + updatedMinutes
                });
            }
        } else {
            this.setState({
                selectedDate: updatedDate
            });
        }

        this.props.onSelect(updatedDate);
    }

    timeSelectCallback(minutes: number) {
        let updatedDate = this.state.selectedDate;
        updatedDate.setHours(Math.floor(minutes / 60));
        updatedDate.setMinutes(minutes % 60);
        this.setState({
            selectedTime: minutes
        });
        this.props.onSelect(updatedDate);
    }

    getMinTimeInMinutes(givenDate: Date) {
        let isSelectedDateToday: boolean = false;
        let today = new Date();
        if (givenDate) {
            isSelectedDateToday =
                givenDate.getDate() == today.getDate() &&
                givenDate.getMonth() == today.getMonth() &&
                givenDate.getFullYear() == today.getFullYear();
        }

        let minTime: number = 0;
        if (isSelectedDateToday) {
            minTime = today.getHours() * 60 + today.getMinutes();
        }
        return minTime;
    }

    static getTimeInMinutes(givenDate: Date) {
        let defaultTime: number = 0;
        if (givenDate) {
            defaultTime = givenDate.getHours() * 60 + givenDate.getMinutes();
        }
        return defaultTime;
    }
}
