// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Calendar, DayOfWeek } from "office-ui-fabric-react/lib/Calendar";
import "./DatePickerView.scss";
import { registerIcons } from "@uifabric/styling";
import { Input, Popup, ChevronStartIcon, ChevronEndIcon, FocusTrapZone, CalendarIcon } from "@fluentui/react-northstar";
import { Constants } from "./../../utils/Constants";
import { UxUtils } from "./../../utils/UxUtils";

registerIcons({
    icons: {
        "chevronLeft": <ChevronStartIcon />,
        "chevronRight": <ChevronEndIcon />
    }
});

let dayPickerStrings = {
    months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    shortMonths: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    shortDays: ["S", "M", "T", "W", "T", "F", "S"],
    prevMonthAriaLabel: "Previous month",
    nextMonthAriaLabel: "Next month",
    goToToday: ""
};

export interface IDatePickerViewState {
    showCalendar: boolean;
    selectedDate: Date | null;
}

export interface IDatePickerViewProps {
    date: Date;
    placeholder?: string;
    minDate?: Date;
    renderForMobile?: boolean;
    disabled?: boolean;
    locale?: string;
    onSelectDate?: (date: Date) => void;
    onDismiss?: () => void;
}

export class DatePickerView extends React.Component<IDatePickerViewProps, IDatePickerViewState> {
    private dateInputRef: HTMLElement;

    public constructor(props: IDatePickerViewProps) {
        super(props);

        this.state = {
            showCalendar: false,
            selectedDate: this.props.date
        };

        this.initializeDate();
    }

    static getDerivedStateFromProps(props, state) {
        return {
            showCalendar: state.showCalendar,
            selectedDate: props.date
        };
    }

    public render(): JSX.Element {
        if (this.props.renderForMobile) {
            return this.renderDatePickerForMobile();
        } else {
            return this.renderDatePickerForWebOrDesktop();
        }
    }

    renderDatePickerForMobile() {
        return (
            <>
                {this.renderDatePickerPreviewView()}
                <input
                    ref={(dateInputRef) => {
                        this.dateInputRef = dateInputRef;
                    }}
                    type="date"
                    aria-label={this.state.selectedDate ? this.state.selectedDate.toLocaleDateString(this.props.locale, {
                        month: "short",
                        day: "numeric",
                        year: "numeric"
                    }) : null}
                    className="hidden-date-input-mob"
                    disabled={this.props.disabled}
                    min={new Date().toISOString().slice(0, 10)}
                    value={this.state.selectedDate ? this.state.selectedDate.toString() : null}
                    onChange={(e) => {
                        if (!this.props.disabled && e.target.value) {
                            this.onDateSelected(new Date(e.target.value));
                        }
                    }}
                    aria-hidden={true}
                />
            </>
        );
    }

    renderDatePickerForWebOrDesktop() {
        return (
            <Popup
                align="start"
                position="below"
                open={!this.props.disabled && this.state.showCalendar}
                onOpenChange={(e, data) => {
                    this.setState((prevState: IDatePickerViewState) => {
                        prevState.showCalendar = data.open;
                        return prevState;
                    });
                }}
                trigger={this.renderDatePickerPreviewView()}
                content={
                    <FocusTrapZone
                        /*
                            This traps the focus within the Calendar component below.
                            On clicking outside the calendar, the calendar is dismissed.
                            Special handling is added for Esc key to dismiss the calendar using keyboard.
                        */
                        onKeyDown={(e) => {
                            if (!e.repeat && (e.keyCode || e.which) == Constants.ESCAPE_ASCII_VALUE && this.state.showCalendar) {
                                this.setState({
                                    showCalendar: false
                                });
                            }
                        }}>
                        <Calendar
                            onSelectDate={(date: Date) => {
                                this.onDateSelected(date);
                            }}
                            isMonthPickerVisible={false}
                            value={this.state.selectedDate}
                            firstDayOfWeek={this.getFirstDayOfWeek()}
                            strings={dayPickerStrings}
                            isDayPickerVisible={true}
                            showGoToToday={false}
                            minDate={this.props.minDate}
                            navigationIcons={{
                                leftNavigation: "chevronLeft",
                                rightNavigation: "chevronRight"
                            }}
                        />
                    </FocusTrapZone>
                }
            />
        );
    }

    renderDatePickerPreviewView() {
        let wrapperClassName = "date-input-view date-picker-preview-container";
        if (this.props.renderForMobile) {
            wrapperClassName += " date-input-view-mob";
        }
        let dateOptions: Intl.DateTimeFormatOptions = { month: "short", day: "2-digit", year: "numeric" };
        if (this.props.disabled) {
            wrapperClassName += " cursor-default";
        }

        let inputWrapperProps = {
            tabIndex: -1,
            "aria-label": (this.props.renderForMobile && this.state.selectedDate) ? this.state.selectedDate.toLocaleDateString(this.props.locale, {
                month: "short",
                day: "numeric",
                year: "numeric"
            }) + ". " + this.props.placeholder : null,
            onClick: () => {
                this.onDatePickerPreviewTap();
            },
            className: wrapperClassName,
            "aria-expanded": this.state.showCalendar,
            ...UxUtils.getTappableInputWrapperRole()
        };

        let inputProps = {
            disabled: this.props.disabled,
            type: "text",
            placeholder: this.props.placeholder,
            "aria-hidden": true,
            value: this.state.selectedDate ? UxUtils.formatDate(this.state.selectedDate, this.props.locale, dateOptions) : null,
            readOnly: true,
            "aria-readonly": false,
            className: "date-input"
        };

        return (
            <Input
                input={{ ...inputProps }}
                wrapper={{ ...inputWrapperProps }}
                icon={<CalendarIcon outline />}
            />
        );
    }

    calendarIconProp() {
        return {
            name: "calendar",
            outline: true,
            className: this.props.disabled ? "cursor-default" : "calendar-icon",
            onClick: () => {
                this.onDatePickerPreviewTap();
            }
        };
    }

    onDatePickerPreviewTap() {
        if (!this.props.disabled) {
            if (this.props.renderForMobile && this.dateInputRef) {
                this.dateInputRef.click();
                this.dateInputRef.focus();
            } else {
                this.setState({
                    showCalendar: !this.state.showCalendar
                });
            }
        }
    }

    onDateSelected(date: Date) {
        if (!this.isValidDate(date)) {
            return;
        }
        if (this.props.onSelectDate) {
            this.props.onSelectDate(date);
        }
        this.setState({
            showCalendar: false,
            selectedDate: date
        });
    }

    isValidDate(date: Date): boolean {
        if (this.props.minDate) {
            if (date.getFullYear() > this.props.minDate.getFullYear()) {
                return true;
            } else if (date.getFullYear() == this.props.minDate.getFullYear()) {
                if (date.getMonth() > this.props.minDate.getMonth()) {
                    return true;
                } else if (date.getMonth() == this.props.minDate.getMonth()) {
                    return (date.getDate() >= this.props.minDate.getDate());
                } else {
                    return false;
                }
            } else {
                return false;
            }
        }
        return true;
    }

    initializeDate(): void {
        // Date for Sunday in Jan month
        let date: Date = new Date("1970-01-04T00:00");
        let locale: string = this.props.locale;

        for (let i = 0; i < 7; i++) {
            dayPickerStrings.days[i] = date.toLocaleDateString(locale, { weekday: "long" });
            dayPickerStrings.shortDays[i] = date.toLocaleDateString(locale, { weekday: "narrow" });
            date.setDate(date.getDate() + 1);
        }

        for (let i = 0; i < 12; i++) {
            dayPickerStrings.months[i] = date.toLocaleDateString(locale, { month: "long" });
            dayPickerStrings.shortMonths[i] = date.toLocaleDateString(locale, { month: "short" });
            date.setMonth(date.getMonth() + 1);
        }
    }

    private getFirstDayOfWeek(): DayOfWeek {
        if (this.props.locale && Constants.LOCALE_TO_FIRST_DAY_OF_WEEK_MAP.hasOwnProperty(this.props.locale.toLowerCase())) {
            return Constants.LOCALE_TO_FIRST_DAY_OF_WEEK_MAP[this.props.locale.toLowerCase()];
        }
        return DayOfWeek.Sunday;
    }

}
