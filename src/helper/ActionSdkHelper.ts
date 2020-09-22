// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as actionSDK from "@microsoft/m365-action-sdk";

export class ActionSdkHelper {

    /*
    * @desc Gets the localized strings in which the app is rendered
    */
    public static async getLocalizedStrings() {
        let request = new actionSDK.GetLocalizedStrings.Request();
        let response = await actionSDK.executeApi(request) as actionSDK.GetLocalizedStrings.Response;
        return response.strings;
    }
    /*
    * @desc Service Request to create new Action Instance
    * @param {actionSDK.Action} action instance which need to get created
    */
    public static async createActionInstance(action: actionSDK.Action) {
        try {
            let createRequest = new actionSDK.CreateAction.Request(action);
            await  actionSDK.executeApi(createRequest);
        } catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    *   @desc Service API Request for getting the actionContext
    *   @return response: {id, error, success, context}
    */
    public static async getContext() {
        try {
            let response = await actionSDK.executeApi(new actionSDK.GetContext.Request()) as actionSDK.GetContext.Response;
            return response;
        } catch(error) {
            console.log("Error: GetContext() "+ error);
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance
    *   @param actionId: actionSDK.ActionSdkContext.actionId
    *   @return response: {id, error, success, action}
    */
    public static async getActionInstance(actionId: string) {
        try {
            let getActionRequest = new actionSDK.GetAction.Request(actionId);
            let response = await actionSDK.executeApi(getActionRequest) as actionSDK.GetAction.Response;
            return {success: true, actionInstance: response};
        } catch(error) {
            console.log("Error: getActionInstance() "+ error);
            return {success: false, error: error};
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance response summary
    *   @param  @param actionId: actionSDK.ActionSdkContext.actionId
    *   @return response: {id, error, success, summary}
    */
    public static async getActionSummary(actionId: string) {
        try {
            let getSummaryRequest = new actionSDK.GetActionDataRowsSummary.Request(actionId, true);
            let response = await actionSDK.executeApi(getSummaryRequest) as actionSDK.GetActionDataRowsSummary.Response;
            return response;
        } catch(error) {
            console.log("Error: getActionSummary() "+ error);
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
    public static async getActionDataRows(context: actionSDK.ActionSdkContext, creatorId=null, continuationToken=null, pageSize=30, dataTableName=null) {
        try {
            let getDataRowsRequest = new actionSDK.GetActionDataRows.Request(context.actionId, creatorId, continuationToken, pageSize, dataTableName);
            let response = await actionSDK.executeApi(getDataRowsRequest) as actionSDK.GetActionDataRows.Response;
            return response;
        } catch(error) {
            console.log("Error: getActionDataRows() "+ error);
        }
    }
    /*
    *   @desc Service API Request for getting the membercount
    *   @param subscription: actionSDK.Subscription
    *   @return response: {id, error, success, memberCount}
    */
    public static async getMemberCount(subscription: actionSDK.Subscription) {
        try {
            let getSubscriptionCount = new actionSDK.GetSubscriptionMemberCount.Request(subscription);
            let response = await actionSDK.executeApi(getSubscriptionCount) as actionSDK.GetSubscriptionMemberCount.Response;
            return response;
        } catch(error) {
            console.log("Error: getMemberCount() "+ error);
        }
    }
    /*
    *   @desc Service API Request for getting the responders details
    *   @param subscription: actionSDK.Subscription
    *   @param userIds: string array of all the datarows creatorId
    *   @return responseResponders: {id, error, success, memberIdsNotFound, members}
    */
    public static async getResponderDetails(subscription: actionSDK.Subscription, userIds: string[]) {
        try {
            let requestResponders = new actionSDK.GetSubscriptionMembers.Request(subscription, userIds);
            let responseResponders = await actionSDK.executeApi(requestResponders) as actionSDK.GetSubscriptionMembers.Response;
            return responseResponders;
        } catch(error) {
            console.log("Error: getResponderDetails() "+ error);
        }
    }

    /*
    *   @desc Service API Request for getting the nonResponders details
    *   @param actionId: context.actionId
    *   @param subscriptionId: context.subscription.id
    *   @return responseNonResponders: {id, error, success, nonParticipantCount, nonParticipants}
    */
    public static async getNonResponders(actionId: string, subscriptionId: string) {
        try {
            let requestNonResponders = new actionSDK.GetActionSubscriptionNonParticipants.Request(actionId, subscriptionId);
            let responseNonResponders = await actionSDK.executeApi(requestNonResponders) as actionSDK.GetActionSubscriptionNonParticipants.Response;
            return  responseNonResponders;
        } catch(error) {
            console.log("Error: getNonResponders() "+ error);
        }
    }

    /*
    *   @desc Service API to delete an action Instance (a survey instance sent to a channel/group)
    *   @param actionId: context.actionId
    *   @return response: {id, error, success}
    */
    public static async deleteActionInstance(actionId) {
        try {
            let request = new actionSDK.DeleteAction.Request(actionId);
            let response = await actionSDK.executeApi(request) as actionSDK.DeleteAction.Response;
            return response;
        } catch(error) {
            console.log("Error: deleteActionInstance() "+ error);
        }

    }

    /*
    *   @desc Service API to download the survey responses in CSV format
    *   @param actionId: context.actionId
    *   @param fileName : name of the CSV file
    *   @return downloadResponse: {id, error, success}
    */
    public static async downloadResponseAsCSV(actionId: string, fileName: string) {
        try {
            let downloadCSVRequest = new actionSDK.DownloadActionDataRowsResult.Request(
                actionId,
                fileName
            );
            let downloadResponse = await actionSDK.executeApi(downloadCSVRequest) as actionSDK.DownloadActionDataRowsResult.Response;
            return downloadResponse;
        } catch(error) {
            console.log("Error: downloadResponseAsCSV() "+ error);
        }
    }

    /*
    *   @desc Service API to Update the status of action Instance (a survey instance sent to a channel/group)
    *   @param updateInfo: object contains the new status for instance, like new due date or survey closed
    *   @return updateActionResponse: {id, error, success}
    */
    public static async updateActionInstanceStatus(updateInfo) {
        try {
            let updateActionRequest = new actionSDK.UpdateAction.Request(updateInfo);
            let updateActionResponse = await actionSDK.executeApi(updateActionRequest) as actionSDK.UpdateAction.Response;
            return updateActionResponse;
        } catch(error) {
            console.log("Error: updateActionInstanceStatus() "+ error);
        }
    }

    /*
    *   @desc Service API to close the adaptive card opened
    */
    public static async closeCardView() {
        try {
            let closeViewRequest = new actionSDK.CloseView.Request();
            await actionSDK.executeApi(closeViewRequest);
        } catch(error) {
            console.log("Error: closeCardView() "+ error);
        }
    }

    /*
    *   @desc Service API to:
    *   1. Add the new response row for action instance (survey sent on channel/group) for a user
    *   2. Add the new response row for the action instance if multiple responses are allowed
    *   3. Update response row if the user has already participated and single response allowed
    */
    public static async addOrUpdateDataRows(addRows, updateRows) {
        try {
            let addOrUpdateRowsRequest = new actionSDK.AddOrUpdateActionDataRows.Request(
                addRows,
                updateRows
            );
            let addOrUpdateResponse = await actionSDK.executeApi(addOrUpdateRowsRequest) as actionSDK.AddOrUpdateActionDataRows.Response;
            return addOrUpdateResponse;
        } catch(error) {
            console.log("Error: addOrUpdateDataRows() "+ error);
        }
    }

    /*
    *   @desc Service API to hide the loader when the data load is successful to show the page or if failed then to show the error
    */
    public static async hideLoadIndicator() {
        await actionSDK.executeApi(new actionSDK.HideLoadingIndicator.Request());
    }
}
