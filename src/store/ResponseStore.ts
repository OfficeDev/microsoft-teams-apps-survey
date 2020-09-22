// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createStore } from "satcheljs";
import "../mutator/ResponseMutator";
import "../orchestrators/ResponseOrchestrators";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ProgressState } from "../utils/SharedEnum";

export enum ResponsePageViewType {
    Main,
    SelectedResponseView,
    MyResponses
}

export enum ResponseViewMode {
    CreationPreview,
    NewResponse,
    UpdateResponse,
    DisabledResponse
}

interface ISurveyResponseStore {
    context: actionSDK.ActionSdkContext;
    actionInstance: actionSDK.Action;
    response: {
        id: string;
        row: {}
    };
    savedActionInstanceRow: {};
    isValidationModeOn: boolean;
    isInitialized: ProgressState;
    isSendActionInProgress: boolean;
    currentView: ResponsePageViewType;
    myResponses: actionSDK.ActionDataRow[];
    currentResponseIndex: number;
    responseViewMode: ResponseViewMode;
    responseSubmissionFailed: boolean;
    topMostErrorIndex: number;
    isActionDeleted: boolean;
}

const store: ISurveyResponseStore = {
    context: null,
    actionInstance: null,
    response: {
        id: null,
        row: {}
    },
    savedActionInstanceRow: {},
    isValidationModeOn: false,
    isInitialized: ProgressState.NotStarted,
    isSendActionInProgress: false,
    currentView: ResponsePageViewType.Main,
    myResponses: [],
    currentResponseIndex: -1,
    responseViewMode: ResponseViewMode.NewResponse,
    responseSubmissionFailed: false,
    topMostErrorIndex: -1,
    isActionDeleted: false
};

export default createStore<ISurveyResponseStore>("responseStore", store);
