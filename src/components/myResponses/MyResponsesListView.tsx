// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { RecyclerViewComponent, RecyclerViewType } from "./../RecyclerViewComponent";
import getStore from "../../store/MyResponseStore";
import { Flex, Text, Divider, Avatar, FlexItem, FocusZone } from "@fluentui/react-northstar";
import { ChevronEndIcon } from "@fluentui/react-icons-northstar";
import { observer } from "mobx-react";
import * as actionSDK from "@microsoft/m365-action-sdk";
import "./MyResponses.scss";
import { Constants } from "./../../utils/Constants";
import { Localizer } from "../../utils/Localizer";
import { Utils } from "../../utils/Utils";
import { UxUtils } from "./../../utils/UxUtils";

export interface IMyResponsesPage {
    onRowClick?: (index, dataSource) => void;
    locale?: string;
}
/**
 * This component renders to display the current user's response for the instance
 * It will show all the response list along with the timestamp of user's response
 */
@observer
export class MyResponsesListView extends React.Component<IMyResponsesPage, any> {
    private responseTimeStamps: string[] = [];

    render() {
        this.responseTimeStamps = [];
        for (let row of getStore().myResponses) {
            this.addUserResponseTimeStamp(row);
        }

        let myUserName: string = Localizer.getString("You");

        return (
            <FocusZone className="zero-padding" isCircularNavigation={true}>
                <Flex column
                    className="list-container"
                    gap="gap.small">
                    <Flex vAlign="center" gap="gap.small">
                        <Flex.Item>
                            <Avatar name={myUserName} size="large" />
                        </Flex.Item>
                        <Flex.Item >
                            <Text content={Localizer.getString("YourResponses(N)", getStore().myResponses.length)} weight="bold" />
                        </Flex.Item>
                    </Flex>
                    <Divider className="divider zero-bottom-margin" />
                    <RecyclerViewComponent
                        data={this.responseTimeStamps}
                        rowHeight={Constants.LIST_VIEW_ROW_HEIGHT}
                        //This will redirect to the user's response at the timestamp specified in the row.
                        onRowRender={(type: RecyclerViewType, index: number, date: string): JSX.Element => {
                            return this.onRowRender(type, index, date);
                        }} />
                </Flex>
            </FocusZone>
        );

    }

    private onRowRender(type: RecyclerViewType, index: number, date: string): JSX.Element {
        return (<>
            <Flex
                vAlign="center"
                className="my-response-item"
                onClick={() => {
                    if(this.props.onRowClick) {
                        this.props.onRowClick(index, getStore().myResponses);
                    }
                }}
                {...UxUtils.getTabKeyProps()} >
                <Text content={date} />
                <FlexItem push>
                    <ChevronEndIcon size="smallest" outline></ChevronEndIcon>
                </FlexItem>
            </Flex>
            <Divider />
        </>);
    }

    private addUserResponseTimeStamp(row: actionSDK.ActionDataRow): void {
        if (row) {
            let responseTimeStamp: string = Utils.dateTimeToLocaleString(new Date(row.updateTime),
                (this.props.locale) ? this.props.locale : Utils.DEFAULT_LOCALE);
            this.responseTimeStamps.push(responseTimeStamp);
        }
    }
}
