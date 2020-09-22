// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { mutator } from "satcheljs";
import getStore from "../store/MyResponseStore";
import { initializeMyResponses } from "../actions/MyResponsesActions";

mutator(initializeMyResponses, (msg) => {
    const store = getStore();
    store.myResponses = msg.actionInstanceRows;
});
