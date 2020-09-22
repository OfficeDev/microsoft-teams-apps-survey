// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { observer } from "mobx-react";
import getStore, { QuestionDrillDownInfo } from "../../store/SummaryStore";
import { Flex, Text, Divider, Avatar, Loader, FlexItem } from "@fluentui/react-northstar";
import { RetryIcon, ChevronStartIcon } from "@fluentui/react-icons-northstar";
import { goBack, fetchActionInstanceRows } from "../../actions/SummaryActions";
import { RecyclerViewComponent, RecyclerViewType } from "./../RecyclerViewComponent";
import { ProgressState } from "./../../utils/SharedEnum";
import { UxUtils } from "./../../utils/UxUtils";
import * as actionSDK from "@microsoft/m365-action-sdk";
import "../Response/Response.scss";
import { QuestionDisplayType } from "../../components/Creation/questionContainer/QuestionDisplayType";
import { Localizer } from "../../utils/Localizer";

export interface IResponseAggregationViewProps {
    questionInfo: QuestionDrillDownInfo;
}

interface IResponseRowProps {
    senderName: string;
    rowData?: any;
    profilePhoto: string;
}

@observer
export default class ResponseAggregationView extends React.Component<IResponseAggregationViewProps, {}> {

    private responseRows = [];
    private threshHoldRow: number = 5;

    componentWillMount() {
        fetchActionInstanceRows();
    }

    render() {
        return (
            <>
                <Flex className="body-container">
                    {this.getBody()}
                </Flex>
                {this.getFooter()}
            </>
        );
    }

    private onRowRender(type: RecyclerViewType, index: number, props: IResponseRowProps, status: ProgressState, rowsFetchCallback): JSX.Element {
        if ((index + this.threshHoldRow) > getStore().actionInstanceRows.length &&
            status !== ProgressState.Failed) {
            rowsFetchCallback();
        }
        if (type === RecyclerViewType.Header) {
            return (
                <>
                    <Text content={this.props.questionInfo.title} className="question-title-response-view" />
                    {this.props.questionInfo.subTitle ? <Text content={this.props.questionInfo.subTitle} /> : null}
                    <Divider className="divider-response-view" />
                </>
            );
        }
        if (type === RecyclerViewType.Footer) {
            if (status === ProgressState.Failed) {
                return (
                    <Flex vAlign="center" hAlign="center" gap="gap.small" {...UxUtils.getTabKeyProps()} onClick={() => {
                        rowsFetchCallback();
                    }}>
                        <Text content={Localizer.getString("ResponseFetchError")}></Text>
                        <RetryIcon />
                    </Flex>
                );
            } else if (status === ProgressState.InProgress) {
                return <Loader className="resp-aggregation-loader" />;
            } else {
                return null;
            }
        }
        switch (this.props.questionInfo.type) {
            case actionSDK.ActionDataColumnValueType.Date:
                return this.rowView(props, index, true /*showResponseInline*/);
            case actionSDK.ActionDataColumnValueType.Text:
            default:
                return this.rowView(props, index, false /*showResponseInline*/);
        }
    }

    private rowView(props: IResponseRowProps, index: number, showResponseInline: boolean): JSX.Element {
        let senderNameView: JSX.Element = <Text content={props.senderName} className="sender-name" />;
        let responseDataView: JSX.Element = props.rowData ? <Text content={props.rowData} className="response-data" /> : null;
        return (
            <>
                <Flex gap="gap.small" className="response-row">
                    <Avatar name={props.senderName} image={props.profilePhoto} className="sender-avatar" aria-hidden={true} />
                    {showResponseInline ?
                        <>
                            <Flex gap="gap.small" className="response-row-container">
                                {senderNameView}
                            </Flex>
                            <FlexItem push>
                                {responseDataView}
                            </FlexItem>
                        </> :
                        <Flex gap="gap.small" className="response-row-container" column>
                            {senderNameView}
                            {responseDataView}
                        </Flex>
                    }
                </Flex>
                {index <= this.responseRows.length - 1 ? <Divider className="divider-response-view" /> : null}
            </>
        );
    }

    private getFooter(): JSX.Element {
        return (
            <Flex className="footer-layout" gap={"gap.smaller"}>
                <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()} onClick={() => {
                    goBack();
                }} >
                    <ChevronStartIcon xSpacing="after" size="small" />
                    <Text content={Localizer.getString("Back")} />
                </Flex>
            </Flex>
        );
    }

    private getBody(): JSX.Element {
        const rowHeight = 100;
        this.responseRows = [];
        for (let row of getStore().actionInstanceRows) {
            this.addRowData(row);
        }
        return (
            <Flex className="response-list-container">
                <RecyclerViewComponent
                    data={this.responseRows}
                    rowHeight={rowHeight}
                    showFooter={getStore().progressStatus.actionInstanceRow.toString()}
                    showHeader="question-title"
                    onRowRender={(type: RecyclerViewType, index: number, props: IResponseRowProps): JSX.Element => {
                        return this.onRowRender(type, index, props, getStore().progressStatus.actionInstanceRow, fetchActionInstanceRows);
                    }}
                    nonDeterministicRendering={true} />
            </Flex>
        );
    }

    private addRowData(row: actionSDK.ActionDataRow): void {
        if (!row || !getStore().actionInstance || !row.columnValues || !row.columnValues[this.props.questionInfo.id]) {
            return;
        }
        //do not add this row when question is single option and particular choice is not selected by user
        //this filtering is done only for single choice and multi choice questions because those responders are needed who have responded to a particular choice
        //otherwise users who have responded to the question will be included in the result set
        if (this.props.questionInfo.type === actionSDK.ActionDataColumnValueType.SingleOption) {
            switch (this.props.questionInfo.displayType) {
                case QuestionDisplayType.FiveNumber:
                case QuestionDisplayType.FiveStar:
                case QuestionDisplayType.TenNumber:
                case QuestionDisplayType.TenStar:
                    if (!(row.columnValues[this.props.questionInfo.id] === (this.props.questionInfo.choiceIndex + 1).toString())) {
                        return;
                    }
                    break;
                default:
                    if (row.columnValues[this.props.questionInfo.id] !== this.props.questionInfo.choiceIndex.toString()) {
                        return;
                    }
            }
        }

        //do not add this row when question is multi option and particular choice is not selected by user
        if (this.props.questionInfo.type === actionSDK.ActionDataColumnValueType.MultiOption) {
            let response: string[] = JSON.parse(row.columnValues[this.props.questionInfo.id]);
            if (response.indexOf(this.props.questionInfo.choiceIndex.toString()) === -1) {
                return;
            }
        }
        let userProfile: actionSDK.SubscriptionMember = getStore().userProfile[row.creatorId];
        let responseRow: Partial<IResponseRowProps> = {};
        if (userProfile) {
            responseRow.senderName = getStore().context.userId == row.creatorId ? Localizer.getString("You") : userProfile.displayName;
            if (userProfile.profilePhotoUrl) {
                responseRow.profilePhoto =  userProfile.profilePhotoUrl;
            }
        } else if (getStore().context.userId === row.creatorId) {
            responseRow.senderName = Localizer.getString("You");
        }
        //Adding this check because only responders need to be shown in this case, not their responses
        if (this.props.questionInfo.type !== actionSDK.ActionDataColumnValueType.SingleOption &&
            this.props.questionInfo.type !== actionSDK.ActionDataColumnValueType.MultiOption) {
            responseRow.rowData = row.columnValues[this.props.questionInfo.id];
        }
        this.responseRows.push(responseRow);
    }

}
