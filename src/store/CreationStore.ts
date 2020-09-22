// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createStore } from "satcheljs";
import "../orchestrators/CreationOrchestrators";
import "../mutator/CreationMutator";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";
import { ISettingsComponentProps } from "../components/Creation/Settings";
import { ProgressState } from "../utils/SharedEnum";

export enum Page {
    Main,
    Settings,
    Preview,
    UpdateQuestion
}

interface ISurveyCreationStore {
    context: actionSDK.ActionSdkContext;
    title: string;
    preview: boolean;
    questions: actionSDK.ActionDataColumn[];
    settings: ISettingsComponentProps;
    activeQuestionIndex: number;
    isValidationModeOn: boolean;
    isInitialized: ProgressState;
    initPending: boolean;
    currentPage: Page;
    previousPage: Page;
    isSendActionInProgress: boolean;
    teamsGroupInitialized: ProgressState;
    draftActionInstanceId: string;
    openChannelPickerDialog: boolean;
    openSettingDialog: boolean;
    isSendSurveyAlertOpen: boolean;
    shouldFocusOnError: boolean;
}

const store: ISurveyCreationStore = {
    context: null,
    title: "",
    preview: false,
    questions: [],
    settings: {
        resultVisibility: actionSDK.Visibility.All,
        dueDate: Utils.getDefaultExpiry(7).getTime(),
        isResponseEditable: true,
        isResponseAnonymous: false,
        isMultiResponseAllowed: false,
        strings: null
    },
    activeQuestionIndex: -1,
    isValidationModeOn: false,
    isInitialized: ProgressState.NotStarted,
    initPending: true,
    currentPage: Page.Main,
    previousPage: Page.Main,
    isSendActionInProgress: false,
    teamsGroupInitialized: ProgressState.NotStarted,
    draftActionInstanceId: "",
    openChannelPickerDialog: false,
    openSettingDialog: false,
    isSendSurveyAlertOpen: false,
    shouldFocusOnError: false
};

export default createStore<ISurveyCreationStore>("store", store);
