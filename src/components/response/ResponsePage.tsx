// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import getStore, { ResponseViewMode } from "../../store/ResponseStore";
import { Flex, Text } from "@fluentui/react-northstar";
import "./Response.scss";
import { observer } from "mobx-react";
import { updateResponse } from "../../actions/ResponseActions";
import * as actionSDK from "@microsoft/m365-action-sdk";
import ResponseView from "./ResponseView";
import { Localizer } from "../../utils/Localizer";
import { UxUtils } from "./../../utils/UxUtils";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

export interface IResponsePage {
    showTitle?: boolean;
    responseViewMode: ResponseViewMode;
}

@observer
export default class ResponsePage extends React.Component<IResponsePage, any> {

    render() {
        ActionSdkHelper.hideLoadIndicator();
        return (
            <Flex gap="gap.smaller" column>
                {UxUtils.renderingForMobile() &&
                    getStore().responseSubmissionFailed &&
                    <Text content={Localizer.getString("ResponseSubmitError")}
                        className="response-error" error />}
                {this.props.showTitle ? <><Text content={getStore().actionInstance.displayName} className="survey-title" /></> : null}
                <ol className={"ol-container"}>
                    {this.questionView()}
                </ol>
            </Flex>
        );
    }

    private questionView(): JSX.Element {
        let questionsView: JSX.Element[] = [];

        getStore().actionInstance.dataTables[0].dataColumns.forEach((column: actionSDK.ActionDataColumn, index: number) => {
            const questionView: JSX.Element = (<ResponseView
                isValidationModeOn={getStore().isValidationModeOn}
                questionNumber={index + 1}
                actionInstanceColumn={column}
                response={getStore().response.row[column.name]}
                callback={(response: any) => {
                    updateResponse(index, response);
                }}
                setErroredFocus={getStore().topMostErrorIndex === index + 1}
                responseState={this.props.responseViewMode}
                locale={getStore().context ? getStore().context.locale : undefined}
            />);
            questionsView.push(<div className="bottom-space">{questionView}</div>);
        });
        return <Flex column>{questionsView}</Flex>;
    }
}
