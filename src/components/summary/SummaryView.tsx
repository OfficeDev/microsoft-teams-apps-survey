// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { observer } from "mobx-react";
import getStore, { SummaryPageViewType, ResponsesListViewType } from "../../store/SummaryStore";
import "./Summary.scss";
import { closeSurvey, surveyCloseAlertOpen, updateDueDate, surveyExpiryChangeAlertOpen, setDueDate, surveyDeleteAlertOpen, deleteSurvey, setCurrentView, downloadCSV, setProgressStatus, setResponseViewType, showResponseView } from "../../actions/SummaryActions";
import { BarChartComponent, IBarChartItem } from "./../BarChartComponent";
import { Flex, Divider, Dialog, Loader, Text, Avatar, ButtonProps, SplitButton } from "@fluentui/react-northstar";
import { MoreIcon, CalendarIcon, BanIcon, TrashCanIcon } from "@fluentui/react-icons-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ResponseAggregationContainer } from "./ResponseAggregationContainer";
import * as html2canvas from "html2canvas";
import { SurveyUtils } from "../../utils/SurveyUtils";
import { Localizer } from "../../utils/Localizer";
import { Utils } from "../../utils/Utils";
import { ProgressState } from "./../../utils/SharedEnum";
import { ErrorView } from "./../ErrorView";
import { AdaptiveMenu, AdaptiveMenuItem, AdaptiveMenuRenderStyle } from "./../Menu";
import { UxUtils } from "./../../utils/UxUtils";
import { Constants } from "./../../utils/Constants";
import { DateTimePickerView } from "../DateTime";

/**
 * This component renders the View user gets when View Result button is clicked
 * getHeaderContainer(): This fucntion consist of the component with survey title, due date and dialog box for change due date, close or delete the survey
 * getTopContainer(): This function consist of the components with participation percentage bar and link to open ResponderView/NonResponderView
 * getShortSummaryContainer(): This function consist of aggregate summary of per survey questions, each summary statement it will redirect you ResponseAggregationView.
 * getFooterView(): FooterView for SummaryView has Download button with Download as CSV and Download as Image option
 */
@observer
export default class SummaryView extends React.Component<any, any> {
    private bodyContainer: React.RefObject<HTMLDivElement>;

    constructor(props) {
        super(props);
        this.bodyContainer = React.createRef();
    }

    render() {
        return (
            <>
                <Flex column className={"body-container no-mobile-footer"} ref={this.bodyContainer} id="bodyContainer">
                    {this.getHeaderContainer()}
                    {this.getTopContainer()}
                    {this.getMyResponseContainer()}
                    {this.getShortSummaryContainer()}
                </Flex>
                {this.getFooterView()}
            </>
        );

    }

    getMenuItems(): AdaptiveMenuItem[] {
        let menuItemList: AdaptiveMenuItem[] = [];
        if (this.isCurrentUserCreator() && this.isSurveyActive()) {
            let changeExpiry: AdaptiveMenuItem = {
                key: "changeDueDate",
                content: Localizer.getString("ChangeDueBy"),
                icon: <CalendarIcon outline={true} />,
                onClick: () => {
                    if (getStore().progressStatus.updateActionInstance != ProgressState.InProgress) {
                        setProgressStatus({ updateActionInstance: ProgressState.NotStarted });
                    }
                    surveyExpiryChangeAlertOpen(true);
                }
            };
            menuItemList.push(changeExpiry);

            let closeSurvey: AdaptiveMenuItem = {
                key: "close",
                content: Localizer.getString("CloseSurvey"),
                icon: <BanIcon outline={true} />,
                onClick: () => {
                    if (getStore().progressStatus.deleteActionInstance != ProgressState.InProgress) {
                        setProgressStatus({ closeActionInstance: ProgressState.NotStarted });
                    }
                    surveyCloseAlertOpen(true);
                }
            };
            menuItemList.push(closeSurvey);
        }

        if (this.isCurrentUserCreator()) {
            let deleteSurvey: AdaptiveMenuItem = {
                key: "delete",
                content: Localizer.getString("DeleteSurvey"),
                icon: <TrashCanIcon outline={true} />,
                onClick: () => {
                    if (getStore().progressStatus.deleteActionInstance != ProgressState.InProgress) {
                        setProgressStatus({ deleteActionInstance: ProgressState.NotStarted });
                    }
                    surveyDeleteAlertOpen(true);
                }
            };
            menuItemList.push(deleteSurvey);
        }
        return menuItemList;
    }

    private getMyResponseContainer(): JSX.Element {
        let myResponseCount = getStore().myRows.length;
        let myProfilePhoto: string;
        let myUserName = Localizer.getString("You");
        let currentUserProfile: actionSDK.SubscriptionMember = getStore().userProfile[getStore().context.userId];
        if (currentUserProfile && currentUserProfile.displayName) {
            myUserName = currentUserProfile.displayName;
        }
        if (currentUserProfile && currentUserProfile.profilePhotoUrl) {
            myProfilePhoto = currentUserProfile.profilePhotoUrl;
        }
        if (myResponseCount > 0) {
            let content = Localizer.getString("YouRespondedNTimes", myResponseCount);
            if (myResponseCount == 1 && getStore().myRows[0].columnValues) {
                content = Localizer.getString("YouResponded");
            }
            return (
                <>
                    <Flex data-html2canvas-ignore="true" className="my-response" gap="gap.small" vAlign="center">
                        <Flex.Item >
                            <Avatar name={myUserName} size="large" image={myProfilePhoto} />
                        </Flex.Item>
                        <Flex.Item >
                            <Text {...UxUtils.getTabKeyProps()} className="underline" weight="regular" color="brand" content={content} onClick={() => {
                                setResponseViewType(ResponsesListViewType.MyResponses);
                                if (myResponseCount === 1) {
                                    showResponseView(0, getStore().myRows);
                                } else {
                                    setCurrentView(SummaryPageViewType.ResponderView);
                                }
                            }} />
                        </Flex.Item>
                    </Flex>
                    <Divider data-html2canvas-ignore="true" />
                </>
            );
        } else {
            return (<>
                <Flex data-html2canvas-ignore="true" className="my-response" gap="gap.small" vAlign="center">
                    <Flex.Item >
                        <Avatar name={myUserName} size="large" image={myProfilePhoto} />
                    </Flex.Item>
                    <Flex.Item >
                        <Text content={Localizer.getString("NotResponded")} />
                    </Flex.Item>
                </Flex>
                <Divider data-html2canvas-ignore="true" />
            </>);
        }
    }

    private getShortSummaryContainer(): JSX.Element {
        return (
            <>
                {this.canCurrentUserViewResults() ?
                    <Flex column>
                        <ResponseAggregationContainer
                            questions={getStore().actionInstance.dataTables[0].dataColumns}
                            responseAggregates={getStore().actionSummary.defaultAggregates}
                            totalResponsesCount={getStore().actionSummary.rowCount} />
                    </Flex>
                    :
                    this.getNonCreatorErrorView()}
            </>
        );

    }

    private getTopContainer(): JSX.Element {
        if (getStore().progressStatus.actionInstance == ProgressState.Completed
            && getStore().progressStatus.memberCount == ProgressState.Completed
            && getStore().progressStatus.actionSummary == ProgressState.Completed) {

            let participationString: string = getStore().actionSummary.rowCount === 1 ?
                Localizer.getString("ParticipationIndicatorSingular", getStore().actionSummary.rowCount, getStore().memberCount)
                : Localizer.getString("ParticipationIndicatorPlural", getStore().actionSummary.rowCount, getStore().memberCount);

            let participationIndicator: JSX.Element;
            if (getStore().actionInstance && getStore().actionInstance.dataTables[0].canUserAddMultipleRows) {
                participationString = (getStore().actionSummary.rowCount === 0)
                    ? Localizer.getString("NoResponse")
                    : (getStore().actionSummary.rowCount === 1)
                        ? Localizer.getString("SingleResponse")
                        : Localizer.getString("XResponsesByYMembers", getStore().actionSummary.rowCount, (getStore().actionSummary.rowCreatorCount));
                participationIndicator = null;
            } else {
                let participationInfoItems: IBarChartItem[] = [];
                let participationPercentage = Math.round((getStore().actionSummary.rowCount / getStore().memberCount) * 100);
                participationInfoItems.push({
                    id: "participation",
                    title: Localizer.getString("Participation", participationPercentage),
                    titleClassName: "participation-title",
                    quantity: getStore().actionSummary.rowCount,
                    hideStatistics: true
                });
                participationIndicator = <BarChartComponent items={participationInfoItems}
                    getBarPercentageString={(percentage: number) => {
                        return Localizer.getString("BarPercentage", percentage);
                    }}
                    totalQuantity={getStore().memberCount} />;
            }

            return (
                <>
                    {participationIndicator}
                    <Flex space="between" className="participation-container">
                        <Flex.Item >
                            {this.canCurrentUserViewResults() ?
                                <Text {...UxUtils.getTabKeyProps()} role="button" className="underline" color="brand" size="small" content={participationString} onClick={() => {
                                    setResponseViewType(ResponsesListViewType.AllResponses);
                                    setCurrentView(SummaryPageViewType.ResponderView);
                                }} /> :
                                <Text content={participationString} />
                            }
                        </Flex.Item>
                    </Flex>
                    <Divider />
                </>
            );
        } else if (getStore().progressStatus.memberCount == ProgressState.Failed
            || getStore().progressStatus.actionSummary == ProgressState.Failed) {
            return <ErrorView
                title={Localizer.getString("GenericError")}
                buttonTitle={Localizer.getString("Close")}
            />;
        } else {
            return <Loader />;
        }
    }

    private getActionInstanceStatusString(): string {
        const options: Intl.DateTimeFormatOptions = { year: "numeric", month: "long", day: "numeric", hour: "numeric", minute: "numeric" };
        if (this.isSurveyActive()) {
            return Localizer.getString("dueByDate",
                UxUtils.formatDate(new Date(getStore().actionInstance.expiryTime),
                    (getStore().context && getStore().context.locale) ? getStore().context.locale : Utils.DEFAULT_LOCALE, options));
        }

        if (getStore().actionInstance.status == actionSDK.ActionStatus.Closed) {
            let expiry: number = getStore().actionInstance.updateTime ? getStore().actionInstance.updateTime : getStore().actionInstance.expiryTime;
            return Localizer.getString("ClosedOn",
                UxUtils.formatDate(new Date(expiry),
                    (getStore().context && getStore().context.locale) ? getStore().context.locale : Utils.DEFAULT_LOCALE, options));
        }

        if (this.isSurveyExpired()) {
            return Localizer.getString("ExpiredOn",
                UxUtils.formatDate(new Date(getStore().actionInstance.expiryTime),
                    (getStore().context && getStore().context.locale) ? getStore().context.locale : Utils.DEFAULT_LOCALE, options));
        }
    }

    private getHeaderContainer(): JSX.Element {

        return (
            <Flex column className="summary-header-container" >
                <Flex vAlign="center" className="title-and-menu-container">
                    <Text className="expiry-status" content={this.getActionInstanceStatusString()} />
                    {this.getMenu()}
                </Flex>
                <Divider className="due-by-label-divider" />
                {this.setupDeleteDialog()}
                {this.setupCloseDialog()}
                {this.setupDuedateDialog()}
            </Flex >
        );
    }

    private getFooterView(): JSX.Element {
        if (getStore().progressStatus.actionInstance != ProgressState.Completed) {
            return null;
        }
        if (UxUtils.renderingForMobile()) {
            return null;
        }

        if (this.canCurrentUserViewResults() === false) {
            return null;
        }

        let content =
            getStore().progressStatus.downloadData == ProgressState.InProgress ? (
                <Loader size="small" />
            ) : (
                    Localizer.getString("Download")
                );

        let menuItems = [];

        menuItems.push(
            this.getDownloadSplitButtonItem("download_image", "DownloadImage")
        );

        menuItems.push(
            this.getDownloadSplitButtonItem("download_responses", "DownloadResponses")
        );

        return menuItems.length > 0 ? (
            <Flex className="footer-layout" gap={"gap.smaller"} hAlign="end">
                <SplitButton
                    key="download_button"
                    id="download"
                    menu={menuItems}
                    button={{
                        content: { content },
                        className: "download-button",
                    }}

                    primary
                    toggleButton={{ "aria-label": "more-options" }}
                    onMainButtonClick={() => this.downloadImage()}
                />
            </Flex>
        ) : null;
    }

    private getDownloadSplitButtonItem(key: string, menuLabel: string): AdaptiveMenuItem {
        let menuItem: AdaptiveMenuItem = {
            key: key,
            className: "break-word",
            content: <Text content={Localizer.getString(menuLabel)} />,
            onClick: () => {
                if (key == "download_image") {
                    this.downloadImage();
                } else if (key == "download_responses") {
                    downloadCSV();
                }
            }
        };
        return menuItem;
    }

    private downloadImage() {
        let bodyContainerDiv = document.getElementById("bodyContainer") as HTMLDivElement;
        let backgroundColorOfResultsImage: string = UxUtils.getBackgroundColorForTheme(getStore().context.theme);
        (html2canvas as any)(bodyContainerDiv, { width: bodyContainerDiv.scrollWidth, height: bodyContainerDiv.scrollHeight, backgroundColor: backgroundColorOfResultsImage }).then((canvas) => {
            let fileName: string = Localizer.getString("SurveyResult", getStore().actionInstance.displayName).substring(0, Constants.ACTION_RESULT_FILE_NAME_MAX_LENGTH) + ".png";
            let base64Image = canvas.toDataURL("image/png");
            if (window.navigator.msSaveBlob) {
                window.navigator.msSaveBlob(canvas.msToBlob(), fileName);
            } else {
                Utils.downloadContent(fileName, base64Image);
            }
        });
    }

    private setupDuedateDialog() {
        return <Dialog
            className="due-date-dialog"
            overlay={{
                className: "dialog-overlay"
            }}
            open={getStore().isChangeExpiryAlertOpen}
            onOpen={(e, { open }) => surveyExpiryChangeAlertOpen(open)}
            cancelButton={SurveyUtils.getDialogButtonProps(Localizer.getString("ChangeDueDate"), Localizer.getString("Cancel"))}
            confirmButton={getStore().progressStatus.updateActionInstance == ProgressState.InProgress ?
                <Loader size="small" /> :
                this.getDueDateDialogConfirmationButtonProps()}
            content={
                <Flex gap="gap.smaller" column>
                    <DateTimePickerView showTimePicker locale={getStore().context.locale} renderForMobile={UxUtils.renderingForMobile()} minDate={new Date()} value={new Date(getStore().dueDate)} placeholderDate={Localizer.getString("SelectADate")} placeholderTime={Localizer.getString("SelectATime")} onSelect={(date: Date) => {
                        setDueDate(date.getTime());
                    }} />
                    {getStore().progressStatus.updateActionInstance == ProgressState.Failed ? <Text content={Localizer.getString("SomethingWentWrong")} error /> : null}
                </Flex>
            }
            header={Localizer.getString("ChangeDueDate")}

            onCancel={() => {
                surveyExpiryChangeAlertOpen(false);
            }}
            onConfirm={() => {
                updateDueDate(getStore().dueDate);
            }}
        />;
    }

    private getDueDateDialogConfirmationButtonProps(): ButtonProps {

        let confirmButtonProps: ButtonProps = {
            // if difference less than 60 secs, keep it disabled
            disabled: Math.abs(getStore().dueDate - getStore().actionInstance.expiryTime) / 1000 <= 60
        };
        Object.assign(confirmButtonProps, SurveyUtils.getDialogButtonProps(Localizer.getString("ChangeDueDate"), Localizer.getString("Change")));
        return confirmButtonProps;
    }

    private getMenu() {
        let menuItems: AdaptiveMenuItem[] = this.getMenuItems();
        if (menuItems.length == 0) {
            return null;
        }
        return (
            <AdaptiveMenu
                className="triple-dot-menu"
                key="survey_options"
                renderAs={UxUtils.renderingForMobile() ? AdaptiveMenuRenderStyle.ACTIONSHEET : AdaptiveMenuRenderStyle.MENU}
                content={<MoreIcon title={Localizer.getString("MoreOptions")} outline aria-hidden={false} role="button" />}
                menuItems={menuItems}
                dismissMenuAriaLabel={Localizer.getString("DismissMenu")}
            />
        );
    }

    private isCurrentUserCreator(): boolean {
        return getStore().actionInstance && getStore().context.userId == getStore().actionInstance.creatorId;
    }

    private isSurveyActive(): boolean {
        return getStore().actionInstance && getStore().actionInstance.status == actionSDK.ActionStatus.Active && !this.isSurveyExpired();
    }

    private canCurrentUserViewResults(): boolean {
        return this.isCurrentUserCreator() || (getStore().actionInstance.dataTables[0].rowsVisibility == actionSDK.Visibility.All);
    }

    private isSurveyExpired(): boolean {
        return getStore().actionInstance.expiryTime < new Date().getTime() || getStore().actionInstance.status == actionSDK.ActionStatus.Expired;
    }

    private setupCloseDialog() {
        return <Dialog
            className="dialog-base"
            overlay={{
                className: "dialog-overlay"
            }}
            open={getStore().isSurveyCloseAlertOpen}
            onOpen={(e, { open }) => surveyCloseAlertOpen(open)}
            cancelButton={SurveyUtils.getDialogButtonProps(Localizer.getString("CloseSurvey"), Localizer.getString("Cancel"))}
            confirmButton={getStore().progressStatus.closeActionInstance == ProgressState.InProgress ?
                <Loader size="small" /> :
                SurveyUtils.getDialogButtonProps(Localizer.getString("CloseSurvey"), Localizer.getString("Confirm"))}
            content={
                <Flex gap="gap.smaller" column>
                    <Text content={Localizer.getString("CloseSurveyConfirmation")} />
                    {getStore().progressStatus.closeActionInstance == ProgressState.Failed ? <Text content={Localizer.getString("SomethingWentWrong")} error /> : null}
                </Flex>
            }
            header={Localizer.getString("CloseSurvey")}
            onCancel={() => {
                surveyCloseAlertOpen(false);
            }}
            onConfirm={() => {
                closeSurvey();
            }}
        />;
    }

    private setupDeleteDialog() {
        return <Dialog
            className="dialog-base"
            overlay={{
                className: "dialog-overlay"
            }}
            open={getStore().isDeleteSurveyAlertOpen}
            onOpen={(e, { open }) => surveyDeleteAlertOpen(open)}
            cancelButton={SurveyUtils.getDialogButtonProps(Localizer.getString("DeleteSurvey"), Localizer.getString("Cancel"))}
            confirmButton={getStore().progressStatus.deleteActionInstance == ProgressState.InProgress ?
                <Loader size="small" /> :
                SurveyUtils.getDialogButtonProps(Localizer.getString("DeleteSurvey"), Localizer.getString("Confirm"))}
            content={
                <Flex gap="gap.smaller" column>
                    <Text content={Localizer.getString("DeleteSurveyConfirmation")} />
                    {getStore().progressStatus.closeActionInstance == ProgressState.Failed ? <Text content="Something went wrong" error /> : null}
                </Flex>}
            header={Localizer.getString("DeleteSurvey")}
            onCancel={() => {
                surveyDeleteAlertOpen(false);
            }
            }
            onConfirm={() => {
                deleteSurvey();
            }}
        />;
    }

    private getNonCreatorErrorView = () => {
        let downloadStr = Localizer.getString("DownloadYourResponse");
        if (getStore().myRows.length > 1) {
            downloadStr = Localizer.getString("DownloadYourResponses");
        }
        return (
            <Flex column className="non-creator-error-image-container">
                 <img src="./images/permission_error.png" className="non-creator-error-image" />
                <Text className="non-creator-error-text">{this.isSurveyActive() ?
                    Localizer.getString("VisibilityCreatorOnlyLabel") : getStore().myRows.length === 0 ? Localizer.getString("NotRespondedLabel")
                        : Localizer.getString("VisibilityCreatorOnlyLabel")}</Text>
                {
                    getStore().myRows.length > 0 ?
                        <a className="download-your-responses-link"
                            onClick={
                                () => { downloadCSV(); }
                            }>
                            {downloadStr}
                        </a> : null
                }
            </Flex>
        );
    }

}
