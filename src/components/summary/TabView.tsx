// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex, Text, Menu } from "@fluentui/react-northstar";
import { ArrowLeftIcon, ChevronStartIcon } from "@fluentui/react-icons-northstar";
import { ResponderView } from "./ResponderView";
import getStore, { SummaryPageViewType, ResponsesListViewType } from "../../store/SummaryStore";
import { NonResponderView } from "./NonResponderView";
import { setCurrentView, goBack } from "../../actions/SummaryActions";
import { INavBarComponentProps, NavBarItemType, NavBarComponent } from "./../NavBarComponent";
import { UxUtils} from "./../../utils/UxUtils";
import { Constants} from "./../../utils/Constants";
import {Localizer} from "../../utils/Localizer";
import { observer } from "mobx-react";

/**
 * This components consist  of the adjacent tabs with  Responder's and NonResponder's list
 * ResponderView: Shows responder's list and each responder row redirects to response of corresponding user
 * NonResponderView: Shows non-responder's list
*/
@observer
export class TabView extends React.Component<any, any> {

    private items = [
        {
            key: "responders",
            role: "tab",
            "aria-selected": getStore().currentView == SummaryPageViewType.ResponderView,
            "aria-label": Localizer.getString("Responders"),
            content: Localizer.getString("Responders"),
            onClick: () => {
                setCurrentView(SummaryPageViewType.ResponderView);
            }
        },
        {
            key: "nonResponders",
            role: "tab",
            "aria-selected": getStore().currentView == SummaryPageViewType.NonResponderView,
            "aria-label": Localizer.getString("NonResponders"),
            content: Localizer.getString("NonResponders"),
            onClick: () => {
                setCurrentView(SummaryPageViewType.NonResponderView);
            }
        }
    ];

    componentDidMount() {
        UxUtils.setFocus(document.body, Constants.FOCUSABLE_ITEMS.All);
    }

    render() {
        let participationString: string = getStore().actionSummary.rowCount === 1 ?
            Localizer.getString("ParticipationIndicatorSingular", getStore().actionSummary.rowCount, getStore().memberCount)
            : Localizer.getString("ParticipationIndicatorPlural", getStore().actionSummary.rowCount, getStore().memberCount);
        if (getStore().actionInstance && getStore().actionInstance.dataTables[0].canUserAddMultipleRows) {
            participationString = (getStore().actionSummary.rowCount === 0)
                ? Localizer.getString("NoResponse")
                : (getStore().actionSummary.rowCount === 1)
                    ? Localizer.getString("SingleResponse")
                    : Localizer.getString("XResponsesByYMembers", getStore().actionSummary.rowCount, (getStore().actionSummary.rowCreatorCount));
        }
        return (

            <Flex column className={"body-container tabview-container no-mobile-footer"}>
                {this.getNavBar()}
                {getStore().responseViewType === ResponsesListViewType.AllResponses &&
                    <>
                        <Text className="participation-title" size="small" weight="bold">{participationString}</Text>
                        <Menu role="tablist" className="tab-view" fluid defaultActiveIndex={0} items={this.items} underlined primary />
                    </>}
                {getStore().currentView == SummaryPageViewType.ResponderView ? <ResponderView /> : <NonResponderView />}

                {this.getFooterElement()}
            </Flex>
        );
    }

    private getFooterElement() {

        if (!UxUtils.renderingForMobile()) {
            return (
                <Flex className="footer-layout tab-view-footer" gap={"gap.smaller"}>
                    <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()} onClick={() => {
                        goBack();
                    }} >
                        <ChevronStartIcon xSpacing="after" size="small" />
                        <Text content={Localizer.getString("Back")} />
                    </Flex>
                </Flex>
            );
        } else {
            return null;
        }
    }

    private getNavBar() {
        if (UxUtils.renderingForMobile()) {
            let navBarComponentProps: INavBarComponentProps = {
                title: Localizer.getString("ViewResponses"),
                leftNavBarItem: {
                    icon: <ArrowLeftIcon size="large" />,
                    ariaLabel: Localizer.getString("Back"),
                    onClick: () => {
                        goBack();
                    },
                    type: NavBarItemType.BACK
                }
            };

            return (
                <NavBarComponent {...navBarComponentProps} />
            );
        } else {
            return null;
        }
    }
}
