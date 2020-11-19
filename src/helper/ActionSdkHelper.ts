// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Logger } from "../utils/Logger";
import * as actionSDK from "@microsoft/m365-action-sdk";

export class ActionSdkHelper {

    /*
    * @desc Gets the localized strings in which the app is rendered
    */
    public static async getLocalizedStrings() {
        let request = new actionSDK.GetLocalizedStrings.Request();
        let response = await actionSDK.executeApi(request) as actionSDK.GetLocalizedStrings.Response;
        if (!response.error) {
            return { success: true, strings: response.strings };
        }
        else {
            Logger.logError(`fetchLocalization failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }
    /*
    * @desc Service Request to create new Action Instance
    * @param {actionSDK.Action} action instance which need to get created
    */
    public static async createActionInstance(action: actionSDK.Action) {
        let request = new actionSDK.CreateAction.Request(action);
        let response = await actionSDK.executeApi(request) as actionSDK.CreateAction.Response;
        if (!response.error) {
            Logger.logInfo(`createActionInstance success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
        }
        else {
            Logger.logError(`createActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }
    /*
    *   @desc Service API Request for getting the actionContext
    *   @return response: {id, error, success, context}
    */
    public static async getContext() {
        let response = await actionSDK.executeApi(new actionSDK.GetContext.Request()) as actionSDK.GetContext.Response;
        if (!response.error) {
            return { success: true, context: response.context };
        }
        else {
            Logger.logError(`GetContext failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance
    *   @param actionId: actionSDK.ActionSdkContext.actionId
    *   @return response: {id, error, success, action}
    */
    public static async getActionInstance(actionId: string) {
        let getActionRequest = new actionSDK.GetAction.Request(actionId);
        let response = await actionSDK.executeApi(getActionRequest) as actionSDK.GetAction.Response;
        if (!response.error) {
            return { success: true, actionInstance: response };
        } else {
            Logger.logError(`GetAction failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance response summary
    *   @param  @param actionId: actionSDK.ActionSdkContext.actionId
    *   @return response: {id, error, success, summary}
    */
    public static async getActionSummary(actionId: string) {
        let getSummaryRequest = new actionSDK.GetActionDataRowsSummary.Request(actionId, true);
        let response = await actionSDK.executeApi(getSummaryRequest) as actionSDK.GetActionDataRowsSummary.Response;
        if (!response.error) {
            return response;
        }
        else {
            Logger.logError(`GetActionDataRowsSummary failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance responses
    *   @param context: actionSDK.ActionSdkContext
    *   @param creatorId: default null
    *   @param continuationToken: defaullt null
    *   @param pageSize: default 30
    *   @param dataTableName: default null
    *   @return response: {id, error, success, dataRows, continuationToken}
    */
    public static async getActionDataRows(context: actionSDK.ActionSdkContext, creatorId = null, continuationToken = null, pageSize = 30, dataTableName = null) {
        let request = new actionSDK.GetActionDataRows.Request(context.actionId, creatorId, continuationToken, pageSize);
        let response = await actionSDK.executeApi(request) as actionSDK.GetActionDataRows.Response;
        if (!response.error) {
            Logger.logInfo(`getActionDataRows success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, dataRows: response.dataRows, continuationToken: response.continuationToken };
        }
        else {
            Logger.logError(`getActionDataRows failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }
    /*
    *   @desc Service API Request for getting the membercount
    *   @param subscription: actionSDK.Subscription
    *   @return response: {id, error, success, memberCount}
    */
    public static async getMemberCount(subscription: actionSDK.Subscription) {
        let request = new actionSDK.GetSubscriptionMemberCount.Request(subscription);
        let response = await actionSDK.executeApi(request) as actionSDK.GetSubscriptionMemberCount.Response;
        if (!response.error) {
            Logger.logInfo(`getSubscriptionMemberCount success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, memberCount: response.memberCount };
        }
        else {
            Logger.logError(`getSubscriptionMemberCount failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }
    /*
    *   @desc Service API Request for getting the responders details
    *   @param subscription: actionSDK.Subscription
    *   @param userIds: string array of all the datarows creatorId
    *   @return responseResponders: {id, error, success, memberIdsNotFound, members}
    */
    public static async getResponderDetails(subscription: actionSDK.Subscription, userIds: string[]) {
        let request = new actionSDK.GetSubscriptionMembers.Request(subscription, userIds);
        let response = await actionSDK.executeApi(request) as actionSDK.GetSubscriptionMembers.Response;
        if (!response.error) {
            return response;
        }
        else {
            Logger.logError(`GetSubscriptionMembers failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }

    /*
    *   @desc Service API Request for getting the nonResponders details
    *   @param actionId: context.actionId
    *   @param subscriptionId: context.subscription.id
    *   @return responseNonResponders: {id, error, success, nonParticipantCount, nonParticipants}
    */
    public static async getNonResponders(actionId: string, subscriptionId: string) {
        let request = new actionSDK.GetActionSubscriptionNonParticipants.Request(actionId, subscriptionId);
        let response = await actionSDK.executeApi(request) as actionSDK.GetActionSubscriptionNonParticipants.Response;
        if (!response.error) {
            return response;
        }
        else {
            Logger.logError(`GetActionSubscriptionNonParticipants failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }

    /*
    *   @desc Service API to delete an action Instance (a survey instance sent to a channel/group)
    *   @param actionId: context.actionId
    *   @return response: {id, error, success}
    */
    public static async deleteActionInstance(actionId) {
        let request = new actionSDK.DeleteAction.Request(actionId);
        let response = await actionSDK.executeApi(request) as actionSDK.DeleteAction.Response;
        if (!response.error) {
            Logger.logInfo(`deleteActionInstance success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, deleteSuccess: response.success };
        } else {
            Logger.logError(`deleteActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    *   @desc Service API to download the survey responses in CSV format
    *   @param actionId: context.actionId
    *   @param fileName : name of the CSV file
    *   @return downloadResponse: {id, error, success}
    */
    public static async downloadResponseAsCSV(actionId: string, fileName: string) {
        let downloadCSVRequest = new actionSDK.DownloadActionDataRowsResult.Request(
            actionId,
            fileName
        );
        let response = await actionSDK.executeApi(downloadCSVRequest) as actionSDK.DownloadActionDataRowsResult.Response;
        if (!response.error) {
            return response;
        }
        else {
            Logger.logError(`deleteActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }

    /*
    *   @desc Service API to Update the status of action Instance (a survey instance sent to a channel/group)
    *   @param updateInfo: object contains the new status for instance, like new due date or survey closed
    *   @return updateActionResponse: {id, error, success}
    */
    public static async updateActionInstanceStatus(updateInfo) {
        let request = new actionSDK.UpdateAction.Request(updateInfo);
        let response = await actionSDK.executeApi(request) as actionSDK.UpdateAction.Response;
        if (!response.error) {
            return { success: true, updateResponse: response };
        }
        else {
            Logger.logError(`UpdateAction failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    *   @desc Service API to close the adaptive card opened
    */
    public static async closeCardView() {
        try {
            let closeViewRequest = new actionSDK.CloseView.Request();
            await actionSDK.executeApi(closeViewRequest);
        } catch (error) {
            Logger.logError("Error: closeCardView() " + error);
        }
    }

    /*
    *   @desc Service API to:
    *   1. Add the new response row for action instance (survey sent on channel/group) for a user
    *   2. Add the new response row for the action instance if multiple responses are allowed
    *   3. Update response row if the user has already participated and single response allowed
    */
    public static async addOrUpdateDataRows(addRows, updateRows) {
        let addOrUpdateRowsRequest = new actionSDK.AddOrUpdateActionDataRows.Request(
            addRows,
            updateRows
        );
        let response = await actionSDK.executeApi(addOrUpdateRowsRequest) as actionSDK.AddOrUpdateActionDataRows.Response;
        if (!response.error) {
            return { success: true, addOrUpdateResponse: response };
        }
        else {
            Logger.logError(`AddOrUpdateActionDataRows failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    *   @desc Service API to hide the loader when the data load is successful to show the page or if failed then to show the error
    */
    public static async hideLoadIndicator() {
        await actionSDK.executeApi(new actionSDK.HideLoadingIndicator.Request());
    }
}
