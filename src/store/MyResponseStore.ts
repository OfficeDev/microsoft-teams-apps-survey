// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createStore } from "satcheljs";
import "../mutator/MyResponsesMutator";
import * as actionSDK from "@microsoft/m365-action-sdk";

interface ISurveyMyResponsesStore {
    myResponses: actionSDK.ActionDataRow[];
    currentActiveIndex: number;
}

const store: ISurveyMyResponsesStore = {
    myResponses: [],
    currentActiveIndex: -1,
};

export default createStore<ISurveyMyResponsesStore>("responsesStore", store);
