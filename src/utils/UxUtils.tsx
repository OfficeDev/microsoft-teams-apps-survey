// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Constants } from "./Constants";

export class UxUtils {
    public static getTabKeyProps() {
        return {
            tabIndex: 0,
            role: "button",
            ...this.getClickOnCarriageReturnHandler()
        };
    }

    public static getTabKeyPropsRoleRadio() {
        return {
            tabIndex: 0,
            role: "radio",
            ...this.getClickOnCarriageReturnHandler()
        };
    }

    public static getListItemProps() {
        return {
            "data-is-focusable": "true",
            ...UxUtils.getClickOnCarriageReturnHandler()
        };
    }

    private static getClickOnCarriageReturnHandler() {
        return {
            onKeyUp: (event: React.KeyboardEvent<HTMLDivElement>) => {
                if ((event.which || event.keyCode) == Constants.CARRIAGE_RETURN_ASCII_VALUE) {
                    (event.currentTarget as HTMLDivElement).click();
                }
            }
        };
    }

    public static getTappableInputWrapperRole() {
        if (this.renderingForiOS()) {
            return {
                role: "combobox"
            };
        }
        return {
            role: "button"
        };
    }

    public static renderingForMobile(): boolean {
        let currentHostClientType = document.body.getAttribute("data-hostclienttype");
        return currentHostClientType && (currentHostClientType == "ios" || currentHostClientType == "android");
    }

    public static renderingForiOS(): boolean {
        let currentHostClientType = document.body.getAttribute("data-hostclienttype");
        return currentHostClientType && (currentHostClientType == "ios");
    }

    public static setFocus(element: HTMLElement, customSelectorTypes: string[]): void {
        if (customSelectorTypes && customSelectorTypes.length > 0 && element) {
            let queryString = customSelectorTypes.join(", ");
            let focusableItem = element.querySelector(queryString);
            if (focusableItem) {
                (focusableItem as HTMLElement).focus();
            }
        }
    }

    public static renderingForAndroid(): boolean {
        let currentHostClientType = document.body.getAttribute("data-hostclienttype");
        return currentHostClientType && (currentHostClientType === "android");
    }

    public static formatDate(selectedDate: Date, locale: string, options?: Intl.DateTimeFormatOptions): string {
        let dateOptions = options ? options : { year: "numeric", month: "long", day: "2-digit", hour: "numeric", minute: "numeric" };
        let formattedDate = selectedDate.toLocaleDateString(locale, dateOptions);
        //check if M01, M02, ...M12 pattern is present in the string, if pattern is present, using numeric representation of the month instead
        if (formattedDate.match(/M[\d]{2}/)) {
            let newOptions = { ...dateOptions, "month": "2-digit" };
            formattedDate = selectedDate.toLocaleDateString(locale, newOptions);
        }
        return formattedDate;
    }

    public static getBackgroundColorForTheme(theme: string): string {
        let backColor: string = Constants.colors.defaultBackgroundColor;
        switch (this.getNonNullString(theme).toLowerCase()) {
            case "dark":
                backColor = Constants.colors.darkBackgroundColor;
                break;
            case "contrast":
                backColor = Constants.colors.contrastBackgroundColor;
                break;
        }
        return backColor;
    }

    public static getNonNullString(str: string): string {
        if (!str) {
            return "";
        } else {
            return str;
        }
    }
}
