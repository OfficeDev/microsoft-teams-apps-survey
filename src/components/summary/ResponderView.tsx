// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { RecyclerViewComponent, RecyclerViewType } from "./../RecyclerViewComponent";
import { fetchActionInstanceRows, showResponseView, fetchMyResponse } from "../../actions/SummaryActions";
import getStore, { ResponsesListViewType } from "../../store/SummaryStore";
import { Loader, Flex, Text, FocusZone } from "@fluentui/react-northstar";
import { RetryIcon } from "@fluentui/react-icons-northstar";
import { observer } from "mobx-react";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../../utils/Localizer";
import { MyResponsesListView } from "../MyResponses/MyResponsesListView";
import { Constants } from "./../../utils/Constants";
import { Utils } from "../../utils/Utils";
import { UserInfoView, IUserInfoViewProps } from "./../UserInfoView";
import { ProgressState } from "./../../utils/SharedEnum";
import { UxUtils } from "./../../utils/UxUtils";

/**
 * It creates the component with responder's list
 * And each reponder row click redirects to corresponding responses
*/
@observer
export class ResponderView extends React.Component<any, any> {

    private threshHoldRow: number = 5;
    private rowsWithUser: IUserInfoViewProps[] = [];

    componentWillMount() {
        if (getStore().responseViewType === ResponsesListViewType.AllResponses && getStore().actionInstanceRows.length == 0) {
            fetchActionInstanceRows();
        } else if (getStore().responseViewType === ResponsesListViewType.MyResponses && getStore().myRows.length == 0) {
            fetchMyResponse();
        }
    }

    render() {
        this.rowsWithUser = [];
        if (getStore().responseViewType === ResponsesListViewType.AllResponses) {
            for (let row of getStore().actionInstanceRows) {
                this.addUserInfoProps(row, false);
            }

            return (
                <FocusZone className="zero-padding" isCircularNavigation={true}>
                    <Flex column
                        className="list-container"
                        gap="gap.small">
                        <RecyclerViewComponent
                            data={this.rowsWithUser}
                            rowHeight={Constants.LIST_VIEW_ROW_HEIGHT}
                            showFooter={getStore().progressStatus.actionInstanceRow.toString()}
                            onRowRender={(type: RecyclerViewType, index: number, props: IUserInfoViewProps): JSX.Element => {
                                return this.onRowRender(type, index, props,
                                    getStore().progressStatus.actionInstanceRow, fetchActionInstanceRows);
                            }} />
                    </Flex>
                </FocusZone>
            );
        } else {
            return (
                <MyResponsesListView
                    locale={getStore().context ? getStore().context.locale : Utils.DEFAULT_LOCALE}
                    onRowClick={(index, dataSource) => {
                        showResponseView(index, dataSource);
                    }} />
            );
        }
    }

    private onRowRender(type: RecyclerViewType, index: number, props: IUserInfoViewProps, status: ProgressState, dataSourceFetchCallback): JSX.Element {
        if ((index + this.threshHoldRow) > getStore().actionInstanceRows.length &&
            status !== ProgressState.Failed) {
            dataSourceFetchCallback();
        }

        if (type == RecyclerViewType.Footer) {
            if (status === ProgressState.Failed) {
                return (
                    <Flex vAlign="center" hAlign="center" gap="gap.small" {...UxUtils.getTabKeyProps()} onClick={() => {
                        dataSourceFetchCallback();
                    }}>
                        <Text content={Localizer.getString("ResponseFetchError")}></Text>
                        <RetryIcon />
                    </Flex>
                );
            } else if (status === ProgressState.InProgress) {
                return <Loader />;
            }
        } else {
            return <UserInfoView {...props}
                onClick={() => {
                    showResponseView(index, getStore().actionInstanceRows);
                }} />;
        }
    }

    private addUserInfoProps(row: actionSDK.ActionDataRow, showDivider: boolean): void {
        if (!row || !getStore().actionInstance) {
            return;
        }
        let userProfile: actionSDK.SubscriptionMember = getStore().userProfile[row.creatorId];
        let dateOptions: Intl.DateTimeFormatOptions = { year: "numeric", month: "long", day: "numeric", hour: "numeric", minute: "numeric" };
        let userProps: Partial<IUserInfoViewProps> = {
            date: UxUtils.formatDate(new Date(row.updateTime),
                (getStore().context && getStore().context.locale) ? getStore().context.locale : Utils.DEFAULT_LOCALE, dateOptions),
            showBelowDivider: showDivider
        };

        if (userProfile) {
            userProps.userName = userProfile.displayName ? userProfile.displayName : Localizer.getString("UnknownMember");
            userProps.accessibilityLabel = Localizer.getString("ResponderAccessibilityLabel", userProps.userName, userProps.date);
        } else if (getStore().context.userId == row.creatorId) {
            userProps.userName = Localizer.getString("You");
        }

        userProps.userName = userProps.userName ? userProps.userName : "";

        this.rowsWithUser.push(userProps as IUserInfoViewProps);
    }
}
