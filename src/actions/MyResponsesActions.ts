// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { action } from "satcheljs";
import * as actionSDK from "@microsoft/m365-action-sdk";

enum SurveyMyResponsesAction {
    initializeMyResponses = "initializeMyResponses"
}

export let initializeMyResponses = action(SurveyMyResponsesAction.initializeMyResponses, (actionInstanceRows: actionSDK.ActionDataRow[]) => ({
    actionInstanceRows: actionInstanceRows
}));
