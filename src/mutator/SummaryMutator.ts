// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { mutator } from "satcheljs";
import getStore, { SummaryPageViewType, ResponsesListViewType } from "../store/SummaryStore";
import {
    setContext,
    showMoreOptions,
    updateSummary,
    updateActionInstance,
    setDueDate,
    updateUserProfileMap,
    setCurrentView,
    goBack,
    updateNonResponders,
    updateMemberCount,
    updateCurrentResponseIndex,
    addActionInstanceRows,
    setProgressStatus,
    updateMyRows,
    surveyCloseAlertOpen,
    surveyDeleteAlertOpen,
    surveyExpiryChangeAlertOpen,
    updateContinuationToken,
    setResponseViewType,
    setSelectedQuestionDrillDownInfo,
    setIsActionDeleted
} from "../actions/SummaryActions";
import { Utils } from "./../utils/Utils";

mutator(setContext, (msg) => {
    const store = getStore();
    store.context = msg.context;
});

mutator(setProgressStatus, (msg) => {
    const store = getStore();
    store.progressStatus = {
        ...getStore().progressStatus,
        ...msg.status
    };
});

mutator(updateMyRows, (msg) => {
    const store = getStore();
    store.myRows = msg.rows;
});

mutator(setDueDate, (msg) => {
    const store = getStore();
    store.dueDate = msg.date;
});

mutator(showMoreOptions, (msg) => {
    const store = getStore();
    store.showMoreOptionsList = msg.showMoreOptions;
});

mutator(updateSummary, (msg) => {
    const store = getStore();
    store.actionSummary = msg.actionInstanceSummary;
});

mutator(updateActionInstance, (msg) => {
    const store = getStore();
    store.actionInstance = msg.actionInstance;
    store.dueDate = msg.actionInstance.expiryTime;
});

mutator(surveyCloseAlertOpen, (msg) => {
    const store = getStore();
    store.isSurveyCloseAlertOpen = msg.open;
});

mutator(surveyExpiryChangeAlertOpen, (msg) => {
    const store = getStore();
    store.isChangeExpiryAlertOpen = msg.open;
});

mutator(surveyDeleteAlertOpen, (msg) => {
    const store = getStore();
    store.isDeleteSurveyAlertOpen = msg.open;
});

mutator(updateUserProfileMap, (msg) => {
    const store = getStore();
    store.userProfile = Object.assign(store.userProfile, msg.userProfileMap);
});

mutator(setCurrentView, (msg) => {
    const store = getStore();
    store.currentView = msg.view;
});

mutator(goBack, () => {
    const store = getStore();
    let currentView: SummaryPageViewType = store.currentView;

    switch (currentView) {
        case SummaryPageViewType.ResponseAggregationView:
        case SummaryPageViewType.ResponderView:
            store.currentView = SummaryPageViewType.Main;
            break;

        case SummaryPageViewType.ResponseView:
            if (store.responseViewType === ResponsesListViewType.MyResponses && store.myRows.length > 0) {
                store.currentView = SummaryPageViewType.Main;
                break;
            }
            store.currentView = SummaryPageViewType.ResponderView;
            break;

        case SummaryPageViewType.NonResponderView:
            store.currentView = SummaryPageViewType.Main;
            break;

        default:
            break;
    }
});

mutator(updateNonResponders, (msg) => {
    const store = getStore();
    const nonResponderList = msg.nonResponder;
    if (!Utils.isEmptyObject(nonResponderList) && nonResponderList.length > 0) {
        nonResponderList.sort((object1, object2) => {
            if (object1.displayName < object2.displayName) {
                return -1;
            }
            if (object1.displayName > object2.displayName) {
                return 1;
            }
            return 0;
        });
    }
    store.nonResponders = msg.nonResponder;
});

mutator(updateMemberCount, (msg) => {
    const store = getStore();
    store.memberCount = msg.memberCount;
});

mutator(updateCurrentResponseIndex, (msg) => {
    const store = getStore();
    store.currentResponseIndex = msg.index;
});

mutator(addActionInstanceRows, (msg) => {
    const store = getStore();
    store.actionInstanceRows = store.actionInstanceRows.concat(msg.rows);
});

mutator(updateContinuationToken, (msg) => {
    const store = getStore();
    store.continuationToken = msg.token;
});

mutator(setResponseViewType, (msg) => {
    const store = getStore();
    store.responseViewType = msg.responseViewType;
});

mutator(setSelectedQuestionDrillDownInfo, (msg) => {
    const store = getStore();
    store.selectedQuestionDrillDownInfo = msg.questionDrillDownInfo;
});

mutator(setIsActionDeleted, (msg) => {
    const store = getStore();
    store.isActionDeleted = msg.isActionDeleted;
});
