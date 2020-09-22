// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { action } from "satcheljs";
import { QuestionDisplayType } from "../components/Creation/questionContainer/QuestionDisplayType";
import { ISettingsComponentProps } from "../components/Creation/Settings";
import { ProgressState } from "./../utils/SharedEnum";
import { Page } from "../store/CreationStore";
import * as actionSDK from "@microsoft/m365-action-sdk";

export enum SurveyCreationAction {
    initialize = "initialize",
    setContext = "setContext",
    addQuestion = "addQuestion",
    deleteQuestion = "deleteQuestion",
    updateQuestion = "updateQuestion",
    updateTitle = "updateTitle",
    moveQuestionUp = "moveQuestionUp",
    moveQuestionDown = "moveQuestionDown",
    updateActiveQuestionIndex = "updateActiveQuestionIndex",
    updateSettings = "updateSettings",
    sendAction = "sendAction",
    previewAction = "previewAction",  // To update response view store
    showPreview = "showPreview", // To show preview screen
    setValidationMode = "setValidationMode",
    setAppInitialized = "setAppInitialized",
    duplicateQuestion = "duplicateQuestion",
    goToPage = "goToPage",
    updateCustomProps = "updateCustomProps",
    setSendingFlag = "setSendingFlag",
    updateChoiceText = "updateChoiceText",
    showUpdateQuestionPage = "showUpdateQuestionPage",
    initializeExternal = "initializeExternal",
    setTeamsGroupInitializationState = "setTeamsGroupInitializationState",
    getTeamsGroupAndChannels = "getTeamsGroupAndChannels",
    updateTeamsGroupAndChannels = "updateTeamsGroupAndChannels",
    resetSurveyToDefault = "resetSurveyToDefault",
    setChannelPickerDialogOpen = "setDialogOpen",
    setSettingDialogOpen = "setSettingDialogOpen",
    setSendSurveyAlertOpen = "setSendSurveyAlertOpen",
    validateAndSend = "validateAndSend",
    setPreviousPage = "setPreviousPage",
    setShouldFocusOnError = "setShouldFocusOnError",
    fetchCurrentContext = "fetchCurrentContext"
}

export let sendAction = action(SurveyCreationAction.sendAction);

export let fetchCurrentContext = action(SurveyCreationAction.fetchCurrentContext);

export let previewAction = action(SurveyCreationAction.previewAction);
export let showPreview = action(SurveyCreationAction.showPreview, (showPreview: boolean) => ({ showPreview: showPreview }));

export let addQuestion = action(SurveyCreationAction.addQuestion, (questionType: actionSDK.ActionDataColumnValueType, displayType: QuestionDisplayType, customProps?: any, renderingForMobile?: boolean) => ({
    questionType: questionType,
    displayType: displayType,
    customProps: customProps,
    renderingForMobile: renderingForMobile
}));

export let deleteQuestion = action(SurveyCreationAction.deleteQuestion, (index: number) => ({ index: index }));

export let updateSettings = action(SurveyCreationAction.updateSettings, (props: ISettingsComponentProps) => ({
    settingProps: props
}));

export let updateQuestion = action(SurveyCreationAction.updateQuestion, (index: number, question: actionSDK.ActionDataColumn) => ({
    questionIndex: index,
    question: question
}));

export let moveQuestionUp = action(SurveyCreationAction.moveQuestionUp, (index: number) => ({
    index: index
}));

export let moveQuestionDown = action(SurveyCreationAction.moveQuestionDown, (index: number) => ({
    index: index
}));

export let updateTitle = action(SurveyCreationAction.updateTitle, (text: string) => ({
    value: text
}));

export let updateActiveQuestionIndex = action(SurveyCreationAction.updateActiveQuestionIndex, (index: number) => ({
    activeIndex: index
}));

export let setValidationMode = action(SurveyCreationAction.setValidationMode, (validationMode: boolean) => ({
    validationMode: validationMode
}));

export let initialize = action(SurveyCreationAction.initialize);

export let setContext = action(SurveyCreationAction.setContext, (context: actionSDK.ActionSdkContext) => ({ context: context }));

export let setAppInitialized = action(SurveyCreationAction.setAppInitialized, (state: ProgressState) => ({ state: state }));

export let duplicateQuestion = action(SurveyCreationAction.duplicateQuestion, (index: number) => ({
    index: index
}));

export let goToPage = action(SurveyCreationAction.goToPage, (page: Page) => ({ page: page }));

export let updateCustomProps = action(SurveyCreationAction.updateCustomProps, (index: number, customProps: any) => ({ index: index, customProps: customProps }));

export let setSendingFlag = action(SurveyCreationAction.setSendingFlag, (value: boolean) => ({ value: value }));

export let updateChoiceText = action(SurveyCreationAction.updateChoiceText, (text: string, choiceIndex: number, questionIndex: number) => ({
    text: text,
    choiceIndex: choiceIndex,
    questionIndex: questionIndex
}));

export let showUpdateQuestionPage = action(SurveyCreationAction.showUpdateQuestionPage, (questionIndex: number) => ({
    questionIndex: questionIndex
}));

export let initializeExternal = action(SurveyCreationAction.initializeExternal, (actionInstance: actionSDK.Action) => ({ actionInstance: actionInstance }));

export let getTeamsGroupAndChannels = action(SurveyCreationAction.getTeamsGroupAndChannels);

export let setTeamsGroupInitializationState = action(SurveyCreationAction.setTeamsGroupInitializationState, (state: ProgressState) => ({
    state: state
}));

export let resetSurveyToDefault = action(SurveyCreationAction.resetSurveyToDefault);

export let setChannelPickerDialogOpen = action(SurveyCreationAction.setChannelPickerDialogOpen, (dialogOpenIndicator: boolean) => ({
    dialogOpenIndicator: dialogOpenIndicator
}));

export let setSettingDialogOpen = action(SurveyCreationAction.setSettingDialogOpen, (openDialog: boolean) => ({
    openDialog: openDialog
}));

export let setSendSurveyAlertOpen = action(SurveyCreationAction.setSendSurveyAlertOpen, (openDialog: boolean) => ({
    openDialog: openDialog
}));

export let validateAndSend = action(SurveyCreationAction.validateAndSend);

export let setPreviousPage = action(SurveyCreationAction.setPreviousPage, (previousPage: Page) => ({
    previousPage: previousPage
}));

export let setShouldFocusOnError = action(SurveyCreationAction.setShouldFocusOnError, (shouldFocus: boolean) => ({
    value: shouldFocus
}));
