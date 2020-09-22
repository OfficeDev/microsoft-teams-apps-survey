// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { action } from "satcheljs";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ProgressState } from "./../utils/SharedEnum";
import { ResponsePageViewType, ResponseViewMode } from "../store/ResponseStore";

enum SurveyResponseAction {
    initialize = "initialize",
    setActionInstance = "setActionInstance",
    setContext = "setContext",
    initializeExternal = "initializeExternal",
    updateResponse = "updateResponse",
    sendResponse = "sendResponse",
    setValidationModeOn = "setValidationModeOn",
    setAppInitialized = "setAppInitialized",
    resetResponse = "resetResponse",
    setResponseViewMode = "setResponseViewMode",
    setSendingFlag = "setSendingFlag",
    showResponseView = "showResponseView",
    setCurrentView = "setCurrentView",
    setSavedActionInstanceRow = "setSavedActionInstanceRow",
    updateCurrentResponseIndex = "updateCurrentResponseIndex",
    setMyResponses = "setMyResponses",
    setCurrentResponse = "setCurrentResponse",
    setResponseSubmissionFailed = "setResponseSubmissionFailed",
    updateTopMostErrorIndex = "updateTopMostErrorIndex",
    setIsActionDeleted = "setIsActionDeleted"
}

export let initialize = action(SurveyResponseAction.initialize);
export let setActionInstance = action(SurveyResponseAction.setActionInstance, (actionInstance: actionSDK.Action) => ({
    actionInstance: actionInstance
}));
export let setContext = action(SurveyResponseAction.setContext, (context: actionSDK.ActionSdkContext) => ({
    context: context
}));
export let updateResponse = action(SurveyResponseAction.updateResponse, (index: number, response: any) => ({ index: index, response: response }));
export let sendResponse = action(SurveyResponseAction.sendResponse);
export let initializeExternal = action(SurveyResponseAction.initializeExternal, (actionInstance: actionSDK.Action, actionInstanceRow: actionSDK.ActionDataRow) => ({
    actionInstance: actionInstance,
    actionInstanceRow: actionInstanceRow
}));
export let setValidationModeOn = action(SurveyResponseAction.setValidationModeOn);
export let setAppInitialized = action(SurveyResponseAction.setAppInitialized, (state: ProgressState) => ({ state: state }));
export let resetResponse = action(SurveyResponseAction.resetResponse);
export let setResponseViewMode = action(SurveyResponseAction.setResponseViewMode, (responseViewMode: ResponseViewMode) => ({
    responseState: responseViewMode
}));
export let setSendingFlag = action(SurveyResponseAction.setSendingFlag, (value: boolean) => ({ value: value }));
export let showResponseView = action(SurveyResponseAction.showResponseView, (index: number, responses: actionSDK.ActionDataRow[]) => ({
    index: index,
    responses: responses
}));
export let setCurrentView = action(SurveyResponseAction.setCurrentView, (view: ResponsePageViewType) => ({
    view: view
}));
export let setSavedActionInstanceRow = action(SurveyResponseAction.setSavedActionInstanceRow, (actionInstanceRow: any) => ({
    actionInstanceRow: actionInstanceRow
}));
export let updateCurrentResponseIndex = action(SurveyResponseAction.updateCurrentResponseIndex, (index: number) => ({
    index: index
}));
export let setMyResponses = action(SurveyResponseAction.setMyResponses, (actionInstanceRows: actionSDK.ActionDataRow[]) => ({
    actionInstanceRows: actionInstanceRows
}));
export let setCurrentResponse = action(SurveyResponseAction.setCurrentResponse, (response: actionSDK.ActionDataRow) => ({
    response: response
}));

export let setResponseSubmissionFailed = action(SurveyResponseAction.setResponseSubmissionFailed, (value: boolean) => ({
    value: value
}));
export let updateTopMostErrorIndex = action(SurveyResponseAction.updateTopMostErrorIndex, (index: number) => ({
    index: index
}));
export let setIsActionDeleted = action(SurveyResponseAction.setIsActionDeleted, (isActionDeleted: boolean) => ({
    isActionDeleted: isActionDeleted
}));
