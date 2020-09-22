// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { mutator } from "satcheljs";
import getStore from "../store/ResponseStore";
import { setActionInstance, updateResponse, initializeExternal, setValidationModeOn, setAppInitialized, resetResponse, setResponseViewMode, setSendingFlag, setCurrentView, setSavedActionInstanceRow, updateCurrentResponseIndex, setMyResponses, setCurrentResponse, setContext, setResponseSubmissionFailed, updateTopMostErrorIndex, setIsActionDeleted } from "../actions/ResponseActions";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { toJS } from "mobx";

mutator(setAppInitialized, (msg) => {
    const store = getStore();
    store.isInitialized = msg.state;
});

mutator(setContext, (msg) => {
    const store = getStore();
    store.context = msg.context;
});

mutator(setActionInstance, (msg) => {
    const store = getStore();
    store.actionInstance = msg.actionInstance;
});

mutator(initializeExternal, (msg) => {
    const store = getStore();
    store.actionInstance = msg.actionInstance;
    if (msg.actionInstanceRow) {
        store.response.id = msg.actionInstanceRow.id;
        store.response.row = msg.actionInstanceRow.columnValues;
    }
});

mutator(updateResponse, (msg) => {
    const store = getStore();

    let index: number = msg.index;
    let response: any = msg.response;
    const column: actionSDK.ActionDataColumn = store.actionInstance.dataTables[0].dataColumns[index];

    switch (column.valueType) {
        case actionSDK.ActionDataColumnValueType.MultiOption:
            store.response.row[column.name] = JSON.stringify(response as string[]);
            break;

        case actionSDK.ActionDataColumnValueType.SingleOption:
        case actionSDK.ActionDataColumnValueType.Text:
        case actionSDK.ActionDataColumnValueType.Numeric:
        case actionSDK.ActionDataColumnValueType.Date:
            store.response.row[column.name] = response as string;
            break;
    }
});

mutator(setValidationModeOn, (msg) => {
    const store = getStore();
    store.isValidationModeOn = true;
});

mutator(resetResponse, (msg) => {
    const store = getStore();
    store.response.row = toJS(store.savedActionInstanceRow);
});

mutator(setResponseViewMode, (msg) => {
    const store = getStore();
    store.responseViewMode = msg.responseState;
});

mutator(setSendingFlag, (msg) => {
    const store = getStore();
    store.isSendActionInProgress = msg.value;
});

mutator(setCurrentView, (msg) => {
    const store = getStore();
    store.currentView = msg.view;
});

mutator(setSavedActionInstanceRow, (msg) => {
    const store = getStore();
    store.savedActionInstanceRow = msg.actionInstanceRow;
});

mutator(updateCurrentResponseIndex, (msg) => {
    const store = getStore();
    store.currentResponseIndex = msg.index;
});

mutator(updateTopMostErrorIndex, (msg) => {
    const store = getStore();
    store.topMostErrorIndex = msg.index;
});

mutator(setMyResponses, (msg) => {
    const store = getStore();
    store.myResponses = msg.actionInstanceRows;
});

mutator(setCurrentResponse, (msg) => {
    const store = getStore();
    if (msg.response) {
        store.response.id = msg.response.id;
        store.response.row = msg.response.columnValues;
    }
});

mutator(setResponseSubmissionFailed, (msg) => {
    const store = getStore();
    store.responseSubmissionFailed = msg.value;
});

mutator(setIsActionDeleted, (msg) => {
    const store = getStore();
    store.isActionDeleted = msg.isActionDeleted;
});
