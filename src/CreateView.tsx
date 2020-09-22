// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import * as ReactDOM from "react-dom";
import CreationPage from "./components/Creation/CreationPage";
import { initialize } from "./actions/CreationActions";
import { ActionRootView } from "./components/ActionRootView";

initialize();
ReactDOM.render(
    <ActionRootView>
        <CreationPage />
    </ActionRootView>,
    document.getElementById("root"));
