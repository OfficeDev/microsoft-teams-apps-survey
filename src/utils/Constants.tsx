// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { DayOfWeek } from "office-ui-fabric-react/lib/Calendar";

export class Constants {
    // ASCII value for carriage return
    public static readonly CARRIAGE_RETURN_ASCII_VALUE = 13;
    public static readonly ESCAPE_ASCII_VALUE = 27;
    public static readonly LIST_VIEW_ROW_HEIGHT: number = 48;
    public static readonly RESPONSE_LOG_TAG: string = "SurveyResponse";
    public static readonly NAV_BAR_MENUITEM_SUBMIT_RESPONSE_ID: string = "submitresponse";
    public static readonly NAV_BAR_MENUITEM_EDIT_RESPONSE_ID: string = "editresponse";
    public static readonly COVER_IMAGE_PROP_KEY: string = "headerImage";
    public static readonly SURVEY_QUESTION_MAX_LENGTH = 240;
    public static readonly SURVEY_CHOICE_MAX_LENGTH = 360;

    public static readonly ACTION_INSTANCE_INDEFINITE_EXPIRY = -1;

    // some OS doesn't support long filenames, so capping the action's title length to this number
    public static readonly ACTION_RESULT_FILE_NAME_MAX_LENGTH: number = 50;

    public static readonly FOCUSABLE_ITEMS = {
        All: ["a[href]", "area[href]", "input:not([disabled])", "select:not([disabled])", "textarea:not([disabled])", "button:not([disabled])", '[tabindex="0"]'],
        LINK: "a[href]",
        AREA_LINK: "area[href]",
        INPUT: "input:not([disabled])",
        SELECT: "select:not([disabled])",
        TEXTAREA: "textarea:not([disabled])",
        BUTTON: "button:not([disabled])",
        TAB: '[tabindex="0"]'
    };

    // The following is a map of locales to their corresponding first day of the week.
    // This map only contains locales which do not have Sunday as their first day of the week.
    // The source for this data is moment-with-locales.js version 2.24.0
    // Note: The keys in this map should be in lowercase
    public static readonly LOCALE_TO_FIRST_DAY_OF_WEEK_MAP = {
        "af": DayOfWeek.Monday,
        "ar-ly": DayOfWeek.Saturday,
        "ar-ma": DayOfWeek.Saturday,
        "ar-tn": DayOfWeek.Monday,
        "ar": DayOfWeek.Saturday,
        "az": DayOfWeek.Monday,
        "be": DayOfWeek.Monday,
        "bg": DayOfWeek.Monday,
        "bm": DayOfWeek.Monday,
        "br": DayOfWeek.Monday,
        "bs": DayOfWeek.Monday,
        "ca": DayOfWeek.Monday,
        "cs": DayOfWeek.Monday,
        "cv": DayOfWeek.Monday,
        "cy": DayOfWeek.Monday,
        "da": DayOfWeek.Monday,
        "de-at": DayOfWeek.Monday,
        "de-ch": DayOfWeek.Monday,
        "de": DayOfWeek.Monday,
        "el": DayOfWeek.Monday,
        "en-sg": DayOfWeek.Monday,
        "en-au": DayOfWeek.Monday,
        "en-gb": DayOfWeek.Monday,
        "en-ie": DayOfWeek.Monday,
        "en-nz": DayOfWeek.Monday,
        "eo": DayOfWeek.Monday,
        "es-do": DayOfWeek.Monday,
        "es": DayOfWeek.Monday,
        "et": DayOfWeek.Monday,
        "eu": DayOfWeek.Monday,
        "fa": DayOfWeek.Saturday,
        "fi": DayOfWeek.Monday,
        "fo": DayOfWeek.Monday,
        "fr-ch": DayOfWeek.Monday,
        "fr": DayOfWeek.Monday,
        "fy": DayOfWeek.Monday,
        "ga": DayOfWeek.Monday,
        "gd": DayOfWeek.Monday,
        "gl": DayOfWeek.Monday,
        "gom-latn": DayOfWeek.Monday,
        "hr": DayOfWeek.Monday,
        "hu": DayOfWeek.Monday,
        "hy-am": DayOfWeek.Monday,
        "id": DayOfWeek.Monday,
        "is": DayOfWeek.Monday,
        "it-ch": DayOfWeek.Monday,
        "it": DayOfWeek.Monday,
        "jv": DayOfWeek.Monday,
        "ka": DayOfWeek.Monday,
        "kk": DayOfWeek.Monday,
        "km": DayOfWeek.Monday,
        "ku": DayOfWeek.Saturday,
        "ky": DayOfWeek.Monday,
        "lb": DayOfWeek.Monday,
        "lt": DayOfWeek.Monday,
        "lv": DayOfWeek.Monday,
        "me": DayOfWeek.Monday,
        "mi": DayOfWeek.Monday,
        "mk": DayOfWeek.Monday,
        "ms-my": DayOfWeek.Monday,
        "ms": DayOfWeek.Monday,
        "mt": DayOfWeek.Monday,
        "my": DayOfWeek.Monday,
        "nb": DayOfWeek.Monday,
        "nl-be": DayOfWeek.Monday,
        "nl": DayOfWeek.Monday,
        "nn": DayOfWeek.Monday,
        "pl": DayOfWeek.Monday,
        "pt": DayOfWeek.Monday,
        "ro": DayOfWeek.Monday,
        "ru": DayOfWeek.Monday,
        "sd": DayOfWeek.Monday,
        "se": DayOfWeek.Monday,
        "sk": DayOfWeek.Monday,
        "sl": DayOfWeek.Monday,
        "sq": DayOfWeek.Monday,
        "sr-cyrl": DayOfWeek.Monday,
        "sr": DayOfWeek.Monday,
        "ss": DayOfWeek.Monday,
        "sv": DayOfWeek.Monday,
        "sw": DayOfWeek.Monday,
        "tet": DayOfWeek.Monday,
        "tg": DayOfWeek.Monday,
        "tl-ph": DayOfWeek.Monday,
        "tlh": DayOfWeek.Monday,
        "tr": DayOfWeek.Monday,
        "tzl": DayOfWeek.Monday,
        "tzm-latn": DayOfWeek.Saturday,
        "tzm": DayOfWeek.Saturday,
        "ug-cn": DayOfWeek.Monday,
        "uk": DayOfWeek.Monday,
        "ur": DayOfWeek.Monday,
        "uz-latn": DayOfWeek.Monday,
        "uz": DayOfWeek.Monday,
        "vi": DayOfWeek.Monday,
        "x-pseudo": DayOfWeek.Monday,
        "yo": DayOfWeek.Monday,
        "zh-cn": DayOfWeek.Monday
    };

    public static readonly colors = {
        defaultBackgroundColor: "#fff",
        darkBackgroundColor: "#252423",
        contrastBackgroundColor: "black"
    };

}
