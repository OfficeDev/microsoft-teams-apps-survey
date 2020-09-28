// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { observer } from "mobx-react";
import getStore, { SummaryPageViewType, ResponsesListViewType } from "../../store/SummaryStore";
import { showResponseView, setCurrentView } from "../../actions/SummaryActions";
import SummaryView from "./SummaryView";
import { UserResponseView } from "./UserResponseView";
import { Loader } from "@fluentui/react-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../../utils/Localizer";
import { TabView } from "./TabView";
import ResponseAggregationView from "./ResponseAggregationView";
import { Utils } from "../../utils/Utils";
import { ProgressState } from "./../../utils/SharedEnum";
import { ErrorView } from "./../ErrorView";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

/**
 * This component creates the  SummaryPage with different Views
 * SummaryView: first Page user sees when View Result button is clicked
 * TabView: Responder's and NonResponder's tab
 * ResponseAggregationView: Responses per question
*/

@observer
export default class SummaryPage extends React.Component<any, any> {

    render() {
        if (getStore().isActionDeleted) {
            ActionSdkHelper.hideLoadIndicator();
            return <ErrorView
                title={Localizer.getString("SurveyDeletedError")}
                subtitle={Localizer.getString("SurveyDeletedErrorDescription")}
                buttonTitle={Localizer.getString("Close")}
            />;
        }

        if (getStore().progressStatus.actionInstance == ProgressState.Failed
            || getStore().progressStatus.actionSummary == ProgressState.Failed
            || getStore().progressStatus.localizationState == ProgressState.Failed
            || getStore().progressStatus.memberCount == ProgressState.Failed) {
                ActionSdkHelper.hideLoadIndicator();
                return <ErrorView
                title={Localizer.getString("GenericError")}
                buttonTitle={Localizer.getString("Close")}
            />;
        }

        if (getStore().progressStatus.actionInstance != ProgressState.Completed
            || getStore().progressStatus.actionSummary != ProgressState.Completed
            || getStore().progressStatus.localizationState != ProgressState.Completed
            || getStore().progressStatus.memberCount != ProgressState.Completed) {
            return <Loader />;
        }

        return this.getView();
    }

    private getView(): JSX.Element {
        ActionSdkHelper.hideLoadIndicator();
        return this.getPageView();
    }

    private getPageView(): JSX.Element {
        if (getStore().currentView == SummaryPageViewType.Main) {
            return <SummaryView />;
        } else if (getStore().currentView == SummaryPageViewType.ResponderView || getStore().currentView == SummaryPageViewType.NonResponderView) {
            return <TabView />;
        } else if (getStore().currentView === SummaryPageViewType.ResponseAggregationView) {
            return (<ResponseAggregationView questionInfo={getStore().selectedQuestionDrillDownInfo} />);
        } else if (getStore().currentView == SummaryPageViewType.ResponseView) {
            let dataSource: actionSDK.ActionDataRow[] = (getStore().responseViewType === ResponsesListViewType.AllResponses)
                ? getStore().actionInstanceRows : getStore().myRows;
            let goBackToView: SummaryPageViewType = SummaryPageViewType.ResponderView;
            if (getStore().responseViewType === ResponsesListViewType.MyResponses && dataSource.length === 1) {
                goBackToView = SummaryPageViewType.Main;
            }
            return (
                <UserResponseView
                    responses={dataSource}
                    goBack={() => { setCurrentView(goBackToView); }}
                    currentResponseIndex={getStore().currentResponseIndex}
                    showResponseView={showResponseView}
                    userProfiles={getStore().userProfile}
                    locale={getStore().context ? getStore().context.locale : Utils.DEFAULT_LOCALE} />);
        }
    }
}
