// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { orchestrator } from "satcheljs";
import { initialize, setActionInstance, sendResponse, setValidationModeOn, setAppInitialized, setSendingFlag, setCurrentView, setSavedActionInstanceRow, showResponseView, updateCurrentResponseIndex, setMyResponses, setResponseViewMode, setCurrentResponse, setContext, setResponseSubmissionFailed, updateTopMostErrorIndex } from "../actions/ResponseActions";
import getStore, { ResponsePageViewType, ResponseViewMode } from "../store/ResponseStore";
import { toJS } from "mobx";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ProgressState } from "./../utils/SharedEnum";
import { SurveyUtils } from "../utils/SurveyUtils";
import { Localizer } from "../utils/Localizer";
import { ActionModelHelper } from "../helper/ActionModelHelper";
import { Utils } from "../utils/Utils";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

orchestrator(initialize, async () => {
    try {
        let response = await ActionSdkHelper.getContext();
        if (response.success) {
            setContext(response.context);
            let localizer = await Localizer.initialize();
            let actionRows = await fetchActionInstanceNow();
            let myResponse = await fetchMyResponsesNow();

            if (localizer && actionRows.success && myResponse.success) {
                if (!getStore().actionInstance.dataTables[0].canUserAddMultipleRows && getStore().myResponses.length > 0) {
                    setCurrentResponse(getStore().myResponses[0]);
                    setResponseViewMode(ResponseViewMode.DisabledResponse);
                }
                setSavedActionInstanceRow(toJS(getStore().response.row));
                setAppInitialized(ProgressState.Completed);
            } else {
                setAppInitialized(ProgressState.Failed);

            }
        } else {
            setAppInitialized(ProgressState.Failed);
        }
    } catch (error) {
        setAppInitialized(ProgressState.Failed);
    }
});

async function fetchActionInstanceNow() {
    let response = await ActionSdkHelper.getActionInstance(getStore().context.actionId);
    if (response.success && response.actionInstance.success) {
        setActionInstance(response.actionInstance.action);
        return { success: true };
    } else {
        return { success: false };
    }
}

async function fetchMyResponsesNow() {
    let response = await SurveyUtils.fetchMyResponses(getStore().context);
    if (response.success) {
        setMyResponses(response.rows);
        return { success: true };
    } else {
        return response;
    }

}

orchestrator(sendResponse, async () => {
    setValidationModeOn();
    if (getStore().actionInstance && getStore().actionInstance.dataTables[0].dataColumns.length > 0) {
        let columns = toJS(getStore().actionInstance.dataTables[0].dataColumns);
        let row = toJS(getStore().response.row);
        let addRows = [];
        let updateRows = [];

        for (let i = 0; i < columns.length; i++) {
            if (!SurveyUtils.isValidResponse(row[columns[i].name], columns[i].allowNullValue, columns[i].valueType)) {
                updateTopMostErrorIndex(i + 1);
                setSendingFlag(false);
                return;
            }
        }
        if (Utils.isEmptyObject(row)) {
            row = {
                0: null
            };
        }
        let actionInstanceRow: actionSDK.ActionDataRow = {
            id: getStore().response.id ? getStore().response.id : "",
            actionId: getStore().context.actionId,
            columnValues: row
        };

        if (getStore().actionInstance.dataTables[0].canUserAddMultipleRows) {
            actionInstanceRow.id = "";
        }

        setSendingFlag(true);
        setResponseSubmissionFailed(false);
        Utils.announceText(Localizer.getString("SubmittingResponse"));
        ActionModelHelper.prepareActionInstanceRow(actionInstanceRow);

        if (getStore().actionInstance.dataTables[0].canUserAddMultipleRows || !getStore().response.id) {
            addRows.push(actionInstanceRow);
        } else {
            updateRows.push(actionInstanceRow);
        }
        try {
            let addOrUpdate = await ActionSdkHelper.addOrUpdateDataRows(addRows, updateRows);
            setSendingFlag(false);
            if (addOrUpdate.success) {
                Utils.announceText(Localizer.getString("Submitted"));
                await ActionSdkHelper.closeCardView();
            } else {
                setResponseSubmissionFailed(true);
                setSendingFlag(false);
                Utils.announceText(Localizer.getString("Failed"));
            }
        } catch (error) {
            setResponseSubmissionFailed(true);
            setSendingFlag(false);
            Utils.announceText(Localizer.getString("SubmissionFailed"));
        }

    }
});

orchestrator(showResponseView, (msg) => {
    let index: number = msg.index;
    if (index >= 0 && msg.responses && index < msg.responses.length) {
        setActionInstance(getStore().actionInstance);
        setCurrentResponse(msg.responses[index]);
        updateCurrentResponseIndex(index);
        setCurrentView(ResponsePageViewType.SelectedResponseView);
    }
});
