// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { validateAndSend } from "./../actions/CreationActions";
import { orchestrator } from "satcheljs";
import { sendAction, previewAction, setSendingFlag, setValidationMode, initialize, setAppInitialized, goToPage, updateActiveQuestionIndex, setSendSurveyAlertOpen, setShouldFocusOnError, fetchCurrentContext, setContext } from "../actions/CreationActions";
import getStore, { Page } from "../store/CreationStore";
import { UxUtils } from "./../utils/UxUtils";
import { Logger } from "./../utils/Logger";
import { toJS } from "mobx";
import { initializeExternal } from "./../actions/ResponseActions";
import { SurveyUtils } from "../utils/SurveyUtils";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../utils/Localizer";
import { ProgressState } from "./../utils/SharedEnum";
import { Utils } from "../utils/Utils";
import { ActionModelHelper } from "../helper/ActionModelHelper";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

orchestrator(initialize, async () => {
    try {
        await Localizer.initialize();
        setAppInitialized(ProgressState.Completed);
    } catch (error) {
        setAppInitialized(ProgressState.Failed);
    }
});

orchestrator(fetchCurrentContext, async () => {
    let actionContext = await ActionSdkHelper.getContext();
    if(actionContext.success) {
        setContext(actionContext.context as actionSDK.ActionSdkContext);
    }
});

orchestrator(sendAction, async () => {
    setSendingFlag(true);
    let actionInstance = getActionInstance();
    ActionModelHelper.prepareActionInstance(actionInstance, toJS(getStore().context));
    try {
        await ActionSdkHelper.createActionInstance(actionInstance);
    } catch(error) {
        Logger.logError("Error: " + JSON.stringify(error)); //Add error log
    }
});

orchestrator(previewAction, () => {
    const firstInvalidQuestionIndex = SurveyUtils.getFirstInvalidQuestionIndex(getStore().questions);
    const isValid: boolean = isSurveyValid(firstInvalidQuestionIndex);
    if (isValid) {
        initializeExternal(getActionInstance(), null);
        setValidationMode(false);
        goToPage(Page.Preview);
    } else {
        announceValidationError(firstInvalidQuestionIndex);
        updateActiveQuestionIndex(firstInvalidQuestionIndex);
    }
});

let getActionInstance = (): actionSDK.Action => {
    let actionInstance: actionSDK.Action = {
        displayName: getStore().title,
        expiryTime: getStore().settings.dueDate,
        dataTables: [
            {
                name: "",
                dataColumns: toJS(getStore().questions),
                attachments: [],
            },
        ],
    };

    if (getStore().settings.resultVisibility === actionSDK.Visibility.Sender) {
        actionInstance.dataTables[0].rowsVisibility = actionSDK.Visibility.Sender;
    } else {
        actionInstance.dataTables[0].rowsVisibility = actionSDK.Visibility.All;
    }
    actionInstance.dataTables[0].canUserAddMultipleRows = getStore().settings.isMultiResponseAllowed;

    return actionInstance;
};

orchestrator(validateAndSend, () => {
    const firstInvalidQuestionIndex = SurveyUtils.getFirstInvalidQuestionIndex(getStore().questions);
    const isValid: boolean = isSurveyValid(firstInvalidQuestionIndex);
    if (isValid) {
        if (SurveyUtils.areAllQuestionsOptional(getStore().questions)) {
            setSendSurveyAlertOpen(true);
        } else {
            sendAction();
        }
    } else {
        if (!UxUtils.renderingForMobile()) {
            setShouldFocusOnError(true);
        }
        announceValidationError(firstInvalidQuestionIndex);
        updateActiveQuestionIndex(firstInvalidQuestionIndex);
    }
});

function isSurveyValid(firstInvalidQuestionIndex: number) {
    setValidationMode(true);
    if (!SurveyUtils.isEmptyOrNull(getStore().title) && getStore().questions.length > 0 && firstInvalidQuestionIndex === -1) {
        return true;
    }
    return false;
}

function announceValidationError(invalidQuestionIndex: number) {
    const errorCount = SurveyUtils.countErrorsPresent(getStore().title, invalidQuestionIndex, getStore().questions);
    if (errorCount > 1) {
        Utils.announceText(Localizer.getString("MultipleRequiredError", errorCount));
    } else {
        Utils.announceText(Localizer.getString("OneRequiredError"));
    }
}
