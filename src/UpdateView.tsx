// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import * as ReactDOM from "react-dom";
import ResponseRenderer from "./components/Response/ResponseRenderer";
import { initialize } from "./actions/ResponseActions";
import { ActionRootView } from "./components/ActionRootView";

initialize();
ReactDOM.render(
    <ActionRootView>
        <ResponseRenderer />
    </ActionRootView>,
    document.getElementById("root"));
