// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex, FlexItem, Text, Avatar, Divider, ChevronEndMediumIcon } from "@fluentui/react-northstar";
import "./UserInfoView.scss";
import { UxUtils } from "./../../utils/UxUtils";

export interface IUserInfoViewProps {
    userName: string;
    subtitle?: string;
    date?: string;
    accessibilityLabel?: string;
    showBelowDivider?: boolean;
    onClick?: () => void;
}

export class UserInfoView extends React.PureComponent<IUserInfoViewProps> {

    render() {
        return (
            <>
                <Flex aria-label={this.props.accessibilityLabel} className="user-info-view overflow-hidden" vAlign="center" gap="gap.small" onClick={this.props.onClick} {...UxUtils.getListItemProps()}>
                    <Avatar className="user-profile-pic" name={this.props.userName} size="medium" aria-hidden="true" />
                    <Flex column className="overflow-hidden">
                        <Text truncated size="medium">{this.props.userName}</Text>
                        {this.props.subtitle &&
                            <Text truncated size="small">{this.props.subtitle}</Text>
                        }
                    </Flex>
                    {this.props.date && <FlexItem push>
                        <Text className="nowrap date-grey" size="small">{this.props.date}</Text>
                    </FlexItem>}
                    {this.props.onClick &&
                        <ChevronEndMediumIcon size="smallest" outline />
                    }
                </Flex>
                {this.props.showBelowDivider ? <Divider /> : null}
            </>
        );
    }
}
