// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as actionSDK from "@microsoft/m365-action-sdk";
import { toJS } from "mobx";
import { mutator } from "satcheljs";
import { Localizer } from "../utils/Localizer";
import { Utils } from "../utils/Utils";

import {
    addQuestion,
    deleteQuestion,
    duplicateQuestion,
    goToPage,
    moveQuestionDown,
    moveQuestionUp,
    setAppInitialized,
    setSendingFlag,
    setValidationMode,
    showPreview,
    updateActiveQuestionIndex,
    updateCustomProps,
    updateQuestion,
    updateChoiceText,
    updateSettings,
    updateTitle,
    showUpdateQuestionPage,
    setContext,
    initializeExternal,
    resetSurveyToDefault,
    setChannelPickerDialogOpen,
    setSettingDialogOpen,
    setSendSurveyAlertOpen,
    setPreviousPage,
    setShouldFocusOnError
} from "../actions/CreationActions";
import getStore, { Page } from "../store/CreationStore";
import { QuestionDisplayType } from "../components/Creation/questionContainer/QuestionDisplayType";
import { SurveyUtils } from "../utils/SurveyUtils";

/**
* This mutator function calls is to update the states to store the data for the current session
*/
mutator(setAppInitialized, (msg) => {
    const store = getStore();
    store.isInitialized = msg.state;
});

mutator(updateTitle, (msg) => {
    const store = getStore();
    store.title = msg.value;
});

mutator(addQuestion, (msg) => {
    const questionType: actionSDK.ActionDataColumnValueType = msg.questionType;
    const displayType: QuestionDisplayType = msg.displayType;
    const store = getStore();
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    let qID = questions.length;
    let question: actionSDK.ActionDataColumn = {
        name: qID.toString(),
        valueType: questionType,
        displayName: "",
        allowNullValue: true,
        options: []
    };
    if (displayType != null) {
        question.properties = JSON.stringify({ "dt": displayType });
        if (displayType == QuestionDisplayType.Select) {
            let option1: actionSDK.ActionDataColumnOption = {
                name: "0",
                displayName: ""
            };
            let option2: actionSDK.ActionDataColumnOption = {
                name: "1",
                displayName: ""
            };
            question.options.push(option1, option2);
        }
        if (displayType == QuestionDisplayType.FiveStar) {
            question.options = SurveyUtils.getRatingQuestionOptions(QuestionDisplayType.FiveStar);
            question.properties = JSON.stringify({ ...{ "dt": displayType }, ...msg.customProps });
        }
    }
    questions.push(question);
    store.questions = questions;
    store.activeQuestionIndex = qID;
    if (msg.renderingForMobile) {
        store.previousPage = store.currentPage;
        store.currentPage = Page.UpdateQuestion;
    }
});

mutator(deleteQuestion, (msg) => {
    let index: number = msg.index;
    const store = getStore();
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    questions.splice(index, 1);
    for (index; index < questions.length; index++) {
        questions[index].name = index.toString();
    }
    store.questions = questions;
    if (store.currentPage == Page.UpdateQuestion) {
        store.previousPage = store.currentPage;
        store.currentPage = Page.Main;
    }
    if (index === getStore().activeQuestionIndex) {
        store.activeQuestionIndex = -1;
    }
});

mutator(updateSettings, (msg) => {
    const store = getStore();
    store.settings = msg.settingProps;
});

mutator(updateQuestion, (msg) => {
    const store = getStore();
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    questions[msg.questionIndex] = msg.question;
    store.questions = questions;
    store.isValidationModeOn = false;
});

mutator(moveQuestionUp, (msg) => {
    if (msg.index == 0) {
        return;
    }
    const store = getStore();
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    let currentQuestion: actionSDK.ActionDataColumn = copyQuestion(questions[msg.index]);
    let previousQuestion: actionSDK.ActionDataColumn = copyQuestion(questions[msg.index - 1]);
    currentQuestion.name = (msg.index - 1).toString();
    previousQuestion.name = msg.index.toString();
    questions[msg.index] = previousQuestion;
    questions[msg.index - 1] = currentQuestion;
    store.activeQuestionIndex = msg.index - 1;
    store.questions = questions;
});

mutator(moveQuestionDown, (msg) => {
    const store = getStore();
    if (msg.index == store.questions.length - 1) {
        return;
    }
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    let currentQuestion = copyQuestion(questions[msg.index]);
    let nextQuestion = copyQuestion(questions[msg.index + 1]);
    currentQuestion.name = (msg.index + 1).toString();
    nextQuestion.name = msg.index.toString();
    questions[msg.index] = nextQuestion;
    questions[msg.index + 1] = currentQuestion;
    store.activeQuestionIndex = msg.index + 1;
    store.questions = questions;
});

mutator(updateActiveQuestionIndex, (msg) => {
    const store = getStore();
    store.activeQuestionIndex = msg.activeIndex;
});

mutator(showPreview, (msg) => {
    const store = getStore();
    store.preview = msg.showPreview;
});

mutator(setValidationMode, (msg) => {
    const store = getStore();
    store.isValidationModeOn = msg.validationMode;
});

mutator(duplicateQuestion, (msg) => {
    const store = getStore();
    let questions: actionSDK.ActionDataColumn[] = toJS(store.questions);
    let currentQuestionCopy = copyQuestion(questions[msg.index]);
    currentQuestionCopy.name = (msg.index + 1).toString();
    let index = questions.length - 1;
    for (index; index >= msg.index + 1; index--) {
        questions[index + 1] = questions[index];
        questions[index + 1].name = (index + 1).toString();
    }
    questions[msg.index + 1] = currentQuestionCopy;
    store.activeQuestionIndex = msg.index + 1;
    store.questions = questions;
});

/**
* This function gets the context as parameter and initialize the variables accordingly.
* variable lastSessionData is Previous session data, if any.
* This is applicable when the package view is relaunched (Edit scenario in preview screen before sending the survey).
*/
mutator(setContext, (msg) => {
    const store = getStore();
    store.context = msg.context;
    store.initPending = false;
    if (!Utils.isEmptyObject(store.context.lastSessionData)) {
        const lastSessionData = store.context.lastSessionData;
        const actionInstance: actionSDK.Action = lastSessionData.action;
        getStore().title = actionInstance.displayName;
        updateQuestions(actionInstance.dataTables[0].dataColumns);
        getStore().settings.resultVisibility = actionInstance.dataTables[0].rowsVisibility;
        getStore().settings.dueDate = actionInstance.expiryTime;
        getStore().settings.isMultiResponseAllowed = actionInstance.dataTables[0].canUserAddMultipleRows;
    }

});

mutator(goToPage, (msg) => {
    const store = getStore();
    store.previousPage = store.currentPage;
    store.currentPage = msg.page;
});

mutator(updateCustomProps, (msg) => {
    const store = getStore();
    const question = store.questions[msg.index];
    question.properties = JSON.stringify(msg.customProps);
    store.questions[msg.index] = question;
});

mutator(setSendingFlag, (msg) => {
    const store = getStore();
    store.isSendActionInProgress = msg.value;
});

mutator(updateChoiceText, (msg) => {
    const store = getStore();
    const questionsCopy = [...store.questions];
    questionsCopy[msg.questionIndex].options[msg.choiceIndex].displayName = msg.text;
    store.questions = questionsCopy;
});

mutator(showUpdateQuestionPage, (msg) => {
    const store = getStore();
    store.activeQuestionIndex = msg.questionIndex;
    store.previousPage = store.currentPage;
    store.currentPage = Page.UpdateQuestion;
});

function copyQuestion(question: actionSDK.ActionDataColumn): actionSDK.ActionDataColumn {
    return { ...question };
}

mutator(initializeExternal, (msg) => {
    const store = getStore();
    store.title = msg.actionInstance.displayName;
    store.questions = msg.actionInstance.dataTables[0].dataColumns;
    store.draftActionInstanceId = msg.actionInstance.id;
});

mutator(resetSurveyToDefault, () => {
    const store = getStore();
    store.previousPage = store.currentPage;
    store.currentPage = Page.Main;
    store.title = "";
    store.questions = [];
});

mutator(setChannelPickerDialogOpen, (msg) => {
    const store = getStore();
    store.openChannelPickerDialog = msg.dialogOpenIndicator;
});

mutator(setSettingDialogOpen, (msg) => {
    const store = getStore();
    store.openSettingDialog = msg.openDialog;
});

mutator(setSendSurveyAlertOpen, (msg) => {
    const store = getStore();
    store.isSendSurveyAlertOpen = msg.openDialog;
});

mutator(setPreviousPage, (msg) => {
    const store = getStore();
    store.previousPage = msg.previousPage;
});

mutator(setShouldFocusOnError, (msg) => {
    const store = getStore();
    store.shouldFocusOnError = msg.value;
});

/**
 * Updates the values in getStore().questions using "cl" i.e columns field in viewData
 * @param questions ~ separated
 */
function updateQuestions(questions: actionSDK.ActionDataColumn[]) {
    let columns: actionSDK.ActionDataColumn[] = getStore().questions;
    let id: number = 0;
    questions.forEach(question => {
        let titleString = question.displayName;
        let column: actionSDK.ActionDataColumn = {
            displayName: titleString,
            name: "",
            // Updating with random value for now. Will be filled with correct value in the preceeding code.
            valueType: question.valueType,
            properties: "",
            options: []
        };
        column.allowNullValue = question.allowNullValue;
        let displayType: number = getDisplayType(question) ;

        let customProperties = getCustomProperty(displayType);
        if (customProperties !== null) {
            column.properties = JSON.stringify(customProperties);
        }

        updateOptions(displayType, column.options, question.options);

        column.name = id.toString();
        id++;
        columns.push(column);
    });
}

function getDisplayType(question: actionSDK.ActionDataColumn) {
    let customProperties = JSON.parse(question.properties);
    if (customProperties && customProperties.hasOwnProperty("dt")) {
        return customProperties.dt;
    }
    return QuestionDisplayType.None;
}

function updateOptions(displayType: QuestionDisplayType, columnOptions: actionSDK.ActionDataColumnOption[], optionsArray: actionSDK.ActionDataColumnOption[]) {
    let optionsCount = 0;
    switch (displayType) {
        case QuestionDisplayType.FiveStar:
        case QuestionDisplayType.FiveNumber:
            optionsCount = 5;
            fillRatingOptions(optionsCount, columnOptions);
            break;
        case QuestionDisplayType.TenStar:
        case QuestionDisplayType.TenNumber:
            optionsCount = 10;
            fillRatingOptions(optionsCount, columnOptions);
            break;
        case QuestionDisplayType.LikeDislike:
            fillLikeDislikeOptions(columnOptions);
            break;
        case QuestionDisplayType.Select:
            fillMultiChoiceOptions(optionsArray, columnOptions);

    }
}

function fillMultiChoiceOptions(optionsArray: actionSDK.ActionDataColumnOption[], columnOptions: actionSDK.ActionDataColumnOption[]) {
    let optionId = 0;
    optionsArray.forEach(option => {
        let columnOption: actionSDK.ActionDataColumnOption = {
            name: optionId.toString(),
            displayName: option.displayName
        };
        optionId++;
        columnOptions.push(columnOption);
    });
}

function fillRatingOptions(count: number, columnOptions: actionSDK.ActionDataColumnOption[]) {
    for (let index = 1; index <= count; index++) {
        let columnOption: actionSDK.ActionDataColumnOption = {
            name: index.toString(),
            displayName: index.toString()
        };
        columnOptions.push(columnOption);
    }
}

function fillLikeDislikeOptions(columnOptions: actionSDK.ActionDataColumnOption[]) {
    let columnOptionLike: actionSDK.ActionDataColumnOption = {
        name: "0",
        displayName: "Like"
    };
    let columnOptionDisLike: actionSDK.ActionDataColumnOption = {
        name: "1",
        displayName: "Dislike"
    };
    columnOptions.push(columnOptionLike);
    columnOptions.push(columnOptionDisLike);
}

function getCustomProperty(displayType: number) {
    switch (displayType) {
        case QuestionDisplayType.Select:
        case QuestionDisplayType.None:
            return {
                dt: displayType
            };
        case QuestionDisplayType.FiveStar:
            return {
                dt: displayType,
                level: 5,
                type: Localizer.getString("StarText")
            };
        case QuestionDisplayType.TenStar:
            return {
                dt: displayType,
                level: 10,
                type: Localizer.getString("StarText")
            };
        case QuestionDisplayType.FiveNumber:
            return {
                dt: displayType,
                level: 5,
                type: Localizer.getString("Number")
            };
        case QuestionDisplayType.TenNumber:
            return {
                dt: displayType,
                level: 10,
                type: Localizer.getString("Number")
            };
        case QuestionDisplayType.LikeDislike:
            return {
                dt: displayType,
                type: Localizer.getString("LikeDislike")
            };
    }
    return null;
}
