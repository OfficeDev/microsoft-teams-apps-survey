// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { RecyclerViewComponent,RecyclerViewType } from "./../RecyclerViewComponent";
import getStore from "../../store/SummaryStore";
import { Flex, Loader, FocusZone } from "@fluentui/react-northstar";
import { observer } from "mobx-react";
import { fetchNonResponders } from "../../actions/SummaryActions";
import { Localizer } from "../../utils/Localizer";
import { UserInfoView, IUserInfoViewProps} from "./../UserInfoView";
import { ProgressState } from "./../../utils/SharedEnum";

/**
 * It creates the component with non-responder's list
*/
@observer
export class NonResponderView extends React.Component {
    componentWillMount() {
        fetchNonResponders();
    }

    render() {
        let rowsWithUser: IUserInfoViewProps[] = [];
        if (getStore().progressStatus.nonResponder == ProgressState.InProgress) {
            return <Loader />;
        }

        if (getStore().progressStatus.nonResponder == ProgressState.Completed) {
            if(getStore().nonResponders) {
            for (let userProfile of getStore().nonResponders) {
                userProfile = getStore().userProfile[userProfile.id];

                if (userProfile) {
                    rowsWithUser.push({
                        userName: userProfile.displayName ? userProfile.displayName : Localizer.getString("UnknownMember"),
                        accessibilityLabel: userProfile.displayName
                    });
                }
            }
        }
        }
        return (
            <FocusZone className="zero-padding" isCircularNavigation={true}>
                <Flex column
                    className="list-container"
                    gap="gap.small">
                    <RecyclerViewComponent
                        data={rowsWithUser}
                        rowHeight={48}
                        onRowRender={(type: RecyclerViewType, index: number, props: IUserInfoViewProps): JSX.Element => {
                            return <UserInfoView {...props} />;
                        }} />
                </Flex>
            </FocusZone>
        );
    }
}
