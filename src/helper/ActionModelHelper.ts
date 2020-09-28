// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";

export namespace ActionModelHelper {
    export function getActionInstanceProperty(
        actionInstance: actionSDK.Action,
        propertyName: string
    ): actionSDK.ActionProperty {
        if (
            actionInstance.customProperties &&
            actionInstance.customProperties.length > 0
        ) {
            for (let property of actionInstance.customProperties) {
                if (property.name == propertyName) {
                    return property;
                }
            }
        }
        return null;
    }

    export function prepareActionInstance(
        actionInstance: actionSDK.Action,
        actionContext: actionSDK.ActionSdkContext
    ) {
        if (Utils.isEmptyString(actionInstance.id)) {
            actionInstance.id = Utils.generateGUID();
            actionInstance.createTime = Date.now();
        }
        actionInstance.updateTime = Date.now();
        actionInstance.creatorId = actionContext.userId;
        actionInstance.actionPackageId = actionContext.actionPackageId;
        actionInstance.version = actionInstance.version || 1;
        actionInstance.dataTables[0].rowsEditable =
            actionInstance.dataTables[0].rowsEditable || true;
        actionInstance.dataTables[0].canUserAddMultipleRows =
            actionInstance.dataTables[0].canUserAddMultipleRows || false;
        actionInstance.dataTables[0].rowsVisibility =
            actionInstance.dataTables[0].rowsVisibility || actionSDK.Visibility.All;
        if (getActionInstanceProperty(actionInstance, "Locale") == null) {
            actionInstance.customProperties = actionInstance.customProperties || [];
            actionInstance.customProperties.push({
                name: "Locale",
                valueType: actionSDK.ActionPropertyValueType.Text,
                value: actionContext.locale,
            });
        }
    }

    export function prepareActionInstanceRow(
        actionInstanceRow: actionSDK.ActionDataRow
    ) {
        if (Utils.isEmptyString(actionInstanceRow.id)) {
            actionInstanceRow.id = Utils.generateGUID();
            actionInstanceRow.createTime = Date.now();
        }
        actionInstanceRow.updateTime = Date.now();
    }

    export function prepareActionInstanceRows(
        actionInstanceRows: actionSDK.ActionDataRow[]
    ) {
        for (let actionInstanceRow of actionInstanceRows) {
            this.prepareActionInstanceRow(actionInstanceRow);
        }
    }
}
