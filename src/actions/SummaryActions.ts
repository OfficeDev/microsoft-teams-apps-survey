// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { action } from "satcheljs";
import { SummaryPageViewType, SummaryProgressStatus, ResponsesListViewType, QuestionDrillDownInfo } from "../store/SummaryStore";
import * as actionSDK from "@microsoft/m365-action-sdk";

enum SurveySummaryAction {
    initialize = "initialize",
    setContext = "setContext",
    updateActionInstance = "updateActionInstance",
    setDueDate = "setDueDate",
    updateSummary = "updateSummary",
    setCurrentView = "setCurrentView",
    showMoreOptions = "showMoreOptions",
    actionInstanceRow = "actionInstanceRow",
    surveyCloseAlertOpen = "surveyCloseAlertOpen",
    surveyExpiryChangeAlertOpen = "surveyExpiryChangeAlertOpen",
    surveyDeleteAlertOpen = "surveyDeleteAlertOpen",
    updateNonResponders = "updateNonResponders",
    updateMemberCount = "updateMemberCount",
    updateProfilePhotos = "updateProfilePhotos",
    updateUserProfileInfo = "updateUserProfileInfo",
    updateMyRows = "updateMyRows",
    setProgressStatus = "setProgressStatus",
    goBack = "goBack",
    fetchActionInstance = "fetchActionInstance",
    fetchUserDetails = "fetchUserDetails",
    actionInstanceSummary = "actionInstanceSummary",
    fetchActionInstanceRows = "fetchActionInstanceRows",
    fetchNonResponders = "fetchNonResponders",
    updateDueDate = "updateDueDate",
    closeSurvey = "closeSurvey",
    deleteSurvey = "deleteSurvey",
    updateContinuationToken = "updateContinuationToken",
    downloadCSV = "downloadCSV",
    fetchLocalization = "fetchLocalization",
    fetchMyResponse = "fetchMyResponse",
    fetchMemberCount = "fetchMemberCount",
    updateCurrentResponseIndex = "updateCurrentResponseIndex",
    showResponseView = "showResponseView",
    setResponseViewType = "setResponseViewType",
    setSelectedQuestionDrillDownInfo = "setSelectedQuestionDrillDownInfo",
    setIsActionDeleted = "setIsActionDeleted"
}

export let initialize = action(SurveySummaryAction.initialize);

export let setContext = action(SurveySummaryAction.setContext, (context: actionSDK.ActionSdkContext) => ({
    context: context
}));

export let fetchUserDetails = action(SurveySummaryAction.fetchUserDetails, (userIds: string[]) => ({
    userIds: userIds
}));

export let fetchActionInstance = action(SurveySummaryAction.fetchActionInstance, ((updateState: boolean) => ({ updateState: updateState })));

export let fetchLocalization = action(SurveySummaryAction.fetchLocalization);

export let fetchMyResponse = action(SurveySummaryAction.fetchMyResponse);

export let fetchMemberCount = action(SurveySummaryAction.fetchMemberCount);

export let fetchActionInstanceSummary = action(SurveySummaryAction.actionInstanceSummary);

export let fetchActionInstanceRows = action(SurveySummaryAction.fetchActionInstanceRows);

export let fetchNonResponders = action(SurveySummaryAction.fetchNonResponders);

export let updateDueDate = action(SurveySummaryAction.updateDueDate, (dueDate: number) => ({
    dueDate: dueDate
}));

export let closeSurvey = action(SurveySummaryAction.closeSurvey);

export let deleteSurvey = action(SurveySummaryAction.deleteSurvey);

export let downloadCSV = action(SurveySummaryAction.downloadCSV);

export let setProgressStatus = action(SurveySummaryAction.setProgressStatus, (status: Partial<SummaryProgressStatus>) => ({
    status: status
}));

export let updateActionInstance = action(SurveySummaryAction.updateActionInstance, (actionInstance: actionSDK.Action) => ({
    actionInstance: actionInstance
}));

export let updateMyRows = action(SurveySummaryAction.updateMyRows, (rows: actionSDK.ActionDataRow[]) => ({
    rows: rows
}));

export let setDueDate = action(SurveySummaryAction.setDueDate, (date: number) => ({
    date: date
}));

export let showMoreOptions = action(SurveySummaryAction.showMoreOptions, (showMoreOptions: boolean) => ({
    showMoreOptions: showMoreOptions
}));

export let setCurrentView = action(SurveySummaryAction.setCurrentView, (view: SummaryPageViewType) => ({
    view: view
}));

export let surveyCloseAlertOpen = action(SurveySummaryAction.surveyCloseAlertOpen, (open: boolean) => ({
    open: open
}));

export let surveyExpiryChangeAlertOpen = action(SurveySummaryAction.surveyExpiryChangeAlertOpen, (open: boolean) => ({
    open: open
}));

export let surveyDeleteAlertOpen = action(SurveySummaryAction.surveyDeleteAlertOpen, (open: boolean) => ({
    open: open
}));

export let addActionInstanceRows = action(SurveySummaryAction.actionInstanceRow, (rows: actionSDK.ActionDataRow[]) => ({
    rows: rows
}));

export let updateSummary = action(SurveySummaryAction.updateSummary, (actionInstanceSummary: actionSDK.ActionDataRowsSummary) => ({
    actionInstanceSummary: actionInstanceSummary
}));

export let updateUserProfileMap = action(SurveySummaryAction.updateUserProfileInfo, (userProfileMap: {}) => ({
    userProfileMap: userProfileMap
}));

export let goBack = action(SurveySummaryAction.goBack);

export let updateNonResponders = action(SurveySummaryAction.updateNonResponders, (nonResponder: actionSDK.SubscriptionMember[]) => ({
    nonResponder: nonResponder
}));

export let updateMemberCount = action(SurveySummaryAction.updateMemberCount, (memberCount: number) => ({
    memberCount: memberCount
}));

export let updateCurrentResponseIndex = action(SurveySummaryAction.updateCurrentResponseIndex, (index: number) => ({
    index: index
}));

export let updateContinuationToken = action(SurveySummaryAction.updateContinuationToken, (token: string) => ({
    token: token
}));

export let showResponseView = action(SurveySummaryAction.showResponseView, (index: number, responses: actionSDK.ActionDataRow[]) => ({
    index: index,
    responses: responses
}));

export let setResponseViewType = action(SurveySummaryAction.setResponseViewType, (responseViewType: ResponsesListViewType) => ({
    responseViewType: responseViewType
}));

export let setSelectedQuestionDrillDownInfo = action(SurveySummaryAction.setSelectedQuestionDrillDownInfo, (questionDrillDownInfo: QuestionDrillDownInfo) => ({
    questionDrillDownInfo: questionDrillDownInfo
}));

export let setIsActionDeleted = action(SurveySummaryAction.setIsActionDeleted, (isActionDeleted: boolean) => ({
    isActionDeleted: isActionDeleted
}));
