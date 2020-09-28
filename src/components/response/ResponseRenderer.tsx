// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import getStore, { ResponsePageViewType, ResponseViewMode } from "../../store/ResponseStore";
import { sendResponse, resetResponse, setResponseViewMode, setCurrentView, setSavedActionInstanceRow, showResponseView, setResponseSubmissionFailed } from "../../actions/ResponseActions";
import { Flex, Button, Text, Loader } from "@fluentui/react-northstar";
import { ChevronStartIcon } from "@fluentui/react-icons-northstar";
import ResponsePage from "./ResponsePage";
import { observer } from "mobx-react";
import { MyResponsesListView } from "../MyResponses/MyResponsesListView";
import { UserResponseView } from "../Summary/UserResponseView";
import { initializeMyResponses } from "../../actions/MyResponsesActions";
import "./Response.scss";
import { Localizer } from  "../../utils/Localizer";
import { Utils } from "../../utils/Utils";
import { UxUtils } from "./../../utils/UxUtils";
import { ProgressState } from "./../../utils/SharedEnum";
import { ErrorView } from "./../ErrorView";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

@observer
export default class ResponseRenderer extends React.Component<any, any> {

    render() {
        if (getStore().isActionDeleted) {
            return <ErrorView
                title={Localizer.getString("SurveyDeletedError")}
                subtitle={Localizer.getString("SurveyDeletedErrorDescription")}
                buttonTitle={Localizer.getString("Close")}
            />;
        }

        if (getStore().isInitialized === ProgressState.NotStarted) {
            return <Loader />;
        } else if (getStore().isInitialized === ProgressState.Failed) {

            return <ErrorView
                title={Localizer.getString("GenericError")}
                buttonTitle={Localizer.getString("Close")}
            />;
        }
        ActionSdkHelper.hideLoadIndicator();
        return this.renderForWebOrDesktop();
    }

    private renderForWebOrDesktop() {
        if (getStore().currentView === ResponsePageViewType.MyResponses) {
            return (
                <>
                    <Flex className="body-container">
                        {this.renderMyResponsesListView()}
                    </Flex>
                    <Flex className="footer-layout" gap={"gap.small"}>
                        <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()} onClick={() => {
                            this.myResponsesViewBackButtonHandler();
                        }} >
                            <ChevronStartIcon xSpacing="after" size="small" />
                            <Text content={Localizer.getString("Back")} />
                        </Flex>
                    </Flex>
                </>
            );
        } else if (getStore().currentView === ResponsePageViewType.SelectedResponseView) {
            return this.renderUserResponseView();
        }
        let shouldShowRespondedNTimesLabel = getStore().actionInstance.dataTables[0].canUserAddMultipleRows && getStore().myResponses.length > 0;
        return (
            <>
                <Flex className="body-container">
                    {this.renderResponsePage()}
                </Flex>
                <Flex className="footer-layout space-between" gap="gap.medium" hAlign="end">
                    <Flex column>
                        {shouldShowRespondedNTimesLabel && this.renderYouRespondedNTimesLabel()}
                        {getStore().responseSubmissionFailed &&
                            <Text content={Localizer.getString("ResponseSubmitError")}
                                className={shouldShowRespondedNTimesLabel ? "response-error" : ""} error />}
                    </Flex>
                    <Flex.Item push>
                        {getStore().responseViewMode === ResponseViewMode.DisabledResponse ?
                            <Button content={Localizer.getString("EditResponse")} primary onClick={() => {
                                /*
                                Any update to this handler should also be made in the NAV_BAR_MENUITEM_EDIT_RESPONSE_ID
                                section in navBarMenuCallback() in ResponseOrchestrator
                                */
                                setResponseViewMode(ResponseViewMode.UpdateResponse);
                            }} /> :
                            <Flex gap="gap.medium">
                                {getStore().responseViewMode === ResponseViewMode.UpdateResponse &&
                                    <Button content={Localizer.getString("Cancel")} onClick={() => {
                                        this.responsePageCancelButtonHandler();
                                    }} />
                                }
                                <Button
                                    primary
                                    loading={getStore().isSendActionInProgress}
                                    disabled={getStore().isSendActionInProgress}
                                    content={getStore().responseViewMode === ResponseViewMode.UpdateResponse ? Localizer.getString("UpdateResponse") : Localizer.getString("SubmitResponse")}
                                    onClick={() => {
                                        /*
                                        Any update to this handler should also be made in the NAV_BAR_MENUITEM_SUBMIT_RESPONSE_ID
                                        section in navBarMenuCallback() in ResponseOrchestrator
                                        */
                                        sendResponse();
                                    }}>
                                </Button>
                            </Flex>
                        }
                    </Flex.Item>
                </Flex>
            </>
        );
    }

    private renderYouRespondedNTimesLabel() {
        return (
            <Flex.Item grow>
                <Text
                    size="small"
                    color="brand"
                    content={getStore().myResponses.length === 1
                        ? Localizer.getString("YouRespondedOnce")
                        : Localizer.getString("YouRespondedNTimes", getStore().myResponses.length)}
                    className="underline" onClick={() => {
                        setSavedActionInstanceRow(getStore().response.row);
                        initializeMyResponses(getStore().myResponses);
                        setCurrentView(ResponsePageViewType.MyResponses);
                    }}
                    {...UxUtils.getTabKeyProps()}
                    aria-label={getStore().myResponses.length === 1
                        ? Localizer.getString("YouRespondedOnce")
                        : Localizer.getString("YouRespondedNTimes", getStore().myResponses.length)}
                />
            </Flex.Item>
        );
    }

    private renderUserResponseView() {
        return (
            <UserResponseView
                responses={getStore().myResponses}
                goBack={() => {
                    setCurrentView(ResponsePageViewType.MyResponses);
                }}
                currentResponseIndex={getStore().currentResponseIndex}
                showResponseView={showResponseView}
                locale={getStore().context ? getStore().context.locale : Utils.DEFAULT_LOCALE} />
        );
    }

    private renderMyResponsesListView() {
        return (
            <MyResponsesListView
                locale={getStore().context ? getStore().context.locale : Utils.DEFAULT_LOCALE}
                onRowClick={(index, dataSource) => {
                    showResponseView(index, dataSource);
                }} />
        );
    }

    private renderResponsePage() {
        return (
            <ResponsePage showTitle responseViewMode={getStore().responseViewMode} />
        );
    }

    private myResponsesViewBackButtonHandler() {
        resetResponse();
        setCurrentView(ResponsePageViewType.Main);
    }

    private responsePageCancelButtonHandler() {
        setResponseSubmissionFailed(false);
        resetResponse();
        setResponseViewMode(ResponseViewMode.DisabledResponse);
    }
}
