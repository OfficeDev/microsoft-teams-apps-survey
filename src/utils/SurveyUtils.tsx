// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { QuestionDisplayType } from "../components/Creation/questionContainer/QuestionDisplayType";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { UxUtils } from "./UxUtils";
import { Utils } from "./Utils";
import { Logger } from "./Logger";
import { ButtonProps } from "@fluentui/react-northstar";
import { Localizer } from "./Localizer";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

/**
* This namespace contains all the generic functions used in survey app
*/
export namespace SurveyUtils {
    export const QUESTION_DIV_ID_PREFIX = "question_div_";
    export const ADDQUESTIONBUTTONID = "add_question_button";

    /**
    * Checks validity of question based on the type and it's required format
    * e.g.: 1. question should be not null 2. for MCQ questions, there should be atleast two non-null options should be provided
    * @param question: actionSDK.ActionDataColumn
    * @return true/false: boolean
    */
    export let isQuestionValid = (question: actionSDK.ActionDataColumn) => {
        if ((!question) || isEmptyOrNull(question.displayName)) {
            return false;
        }
        if (question.valueType == actionSDK.ActionDataColumnValueType.SingleOption
            || question.valueType == actionSDK.ActionDataColumnValueType.MultiOption) {

            if (!question.options || question.options.length < 2) {
                return false;
            }
            for (let i = 0; i < question.options.length; i++) {
                if (isEmptyOrNull(question.options[i].displayName)) {
                    return false;
                }
            }
        }
        return true;
    };

    /**
    * Checks if the provided parameter is empty or null
    * @param value: string
    * @return true/false: boolean
    */
    export let isEmptyOrNull = (value: string) => {
        if (!value || value.trim().length === 0) {
            return true;
        }
        return false;
    };

    /**
    * Rating questions are converted and stored as MCQ type questions
    * And based on the rating range the number of options are created
    * @param ratingType: QuestionDisplayType
    * @return options: []
    */
    export let getRatingQuestionOptions = (ratingType: QuestionDisplayType) => {
        let options = [];
        let maxRatings: number = 5;
        if (ratingType == QuestionDisplayType.LikeDislike) {
            maxRatings = 2;
        } else if (ratingType == QuestionDisplayType.TenNumber ||
            ratingType == QuestionDisplayType.TenStar) {
            maxRatings = 10;
        }
        for (let i = 1; i <= maxRatings; i++) {
            let option = {
                name: i.toString(),
                displayName: i.toString()
            };
            options.push(option);
        }
        if (ratingType == QuestionDisplayType.LikeDislike) {
            options[0].name = "0";
            options[0].displayName = "Like";
            options[1].name = "1";
            options[1].displayName = "Dislike";
        }
        return options;
    };

    /**
    * It fetches the response for the survey of logged-in user
    * @param context: actionSDK.ActionSdkContext
    * @param pageSize: number
    * @param rows:  actionSDK.ActionDataRow[]
    * @param continuationToken: string
    * @return  Promise<actionSDK.ActionDataRow[]>(async(resolve, reject): resolve when call is successfull and have result, and reject for errors
    */
    export async function fetchMyResponses(context: actionSDK.ActionSdkContext, pageSize: number = 100, rows: actionSDK.ActionDataRow[] = [], continuationToken: string = null) {
        try {
            let datarowsCall = await ActionSdkHelper.getActionDataRows(context, "self", continuationToken, pageSize, null);
            if(datarowsCall.success) {
                rows = datarowsCall.dataRows;
                if (datarowsCall.continuationToken) {
                    let response = fetchMyResponses(context, pageSize, rows, datarowsCall.continuationToken);
                    return response;
                } else {
                    return {success: true, rows: rows};
                }
            } else {
                return {success: false, error: datarowsCall.error};
            }
        } catch (error) {
            Logger.logError("Error: " + JSON.stringify(error));
            return {success: false, error: error};
        }
    }

    /**
    * This is a check to validate whether all questions are optional
    * The return value is use to show alert to creator to validate if creator wants to send all optional questions in the survey
    * @param questions: actionSDK.ActionDataColumn[]
    * @return true/false: boolean
    */
    export function areAllQuestionsOptional(questions: actionSDK.ActionDataColumn[]): boolean {
        for (let i = 0; i < questions.length; i++) {
            if (!questions[i].allowNullValue) {
                return false;
            }
        }
        return true;
    }

    /**
    * This check is to get and point the first invalid question in the createPage to announce validation error
    * @param questions: actionSDK.ActionDataColumn[]
    * @return true/false: boolean
    */
    export function getFirstInvalidQuestionIndex(questions: actionSDK.ActionDataColumn[]): number {
        for (let i = 0; i < questions.length; i++) {
            if (!isQuestionValid(questions[i])) {
                return i;
            }
        }
        return -1;
    }

    /**
    * It counts the numbers of validation errors in the page and show the message accordingly
    * @param surveyTitle: string
    * @param firstInvalidQuestionIndex: number
    * @param questions:  actionSDK.ActionDataColumn[]
    * @return numErrors: number
    */
    export function countErrorsPresent(surveyTitle: string, firstInvalidQuestionIndex: number, questions: actionSDK.ActionDataColumn[]): number {
        let numErrors = 0;
        if (isEmptyOrNull(surveyTitle)) {
            numErrors++;
        }
        if (firstInvalidQuestionIndex === -1) {
            return numErrors;
        }
        for (let i = firstInvalidQuestionIndex; i < questions.length; i++) {
            let question = questions[i];
            if (isEmptyOrNull(question.displayName)) {
                numErrors++;
            }
            if (question.valueType == actionSDK.ActionDataColumnValueType.SingleOption
                || question.valueType == actionSDK.ActionDataColumnValueType.MultiOption) {
                for (let i = 0; i < question.options.length; i++) {
                    if (isEmptyOrNull(question.options[i].displayName)) {
                        numErrors++;
                    }
                }
            }
        }
        return numErrors;
    }

    /**
    * It checks if the response for a particular question type is valid or not like for numeric questions response should be a number
    * @param response: any
    * @param isOptional: boolean (if optonal then response can be null or empty)
    * @param columnType:  actionSDK.ActionDataColumnValueType
    * @return true/false: boolean
    */
    export function isValidResponse(response: any, isOptional: boolean, columnType: actionSDK.ActionDataColumnValueType): boolean {
        switch (columnType) {
            case actionSDK.ActionDataColumnValueType.MultiOption:
                return isOptional || (!Utils.isEmptyObject(response) && JSON.parse(response).length > 0);
            case actionSDK.ActionDataColumnValueType.Numeric:
                if (isInvalidNumericPattern(response)) {
                    return false;
                }
                return isOptional || !Utils.isEmptyObject(response);
            default:
                return isOptional || !Utils.isEmptyObject(response);
        }
    }

    /**
    * It verifies if the number is a valid numeric value or not
    * @param response: any
    * @return true/false: boolean
    */
    export function isInvalidNumericPattern(response: any): boolean {
        return !Utils.isEmptyObject(response) && isNaN(parseFloat(response));
    }

    export function getDialogButtonProps(dialogDescription: string, buttonLabel: string): ButtonProps {
        let buttonProps: ButtonProps = {
            "content": buttonLabel
        };

        if (UxUtils.renderingForMobile()) {
            Object.assign(buttonProps, {
                "aria-label": Localizer.getString("DialogTalkback", dialogDescription, buttonLabel),
            });
        }
        return buttonProps;
    }
}
