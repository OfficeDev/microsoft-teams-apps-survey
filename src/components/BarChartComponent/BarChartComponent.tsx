// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "./BarChartComponent.scss";
import { Flex, Text, FlexItem } from "@fluentui/react-northstar";
import { ShimmerContainer } from "../ShimmerLoader";

export interface IBarChartComponentProps {
    items: IBarChartItem[];
    totalQuantity: number;
    accessibilityLabel?: string;
    getBarPercentageString?: (percentage: number) => string;
    showShimmer?: boolean;
    onItemClicked?: (choiceIndex: number) => void;
}

export interface IBarChartItem {
    title: string;
    titleClassName?: string;
    quantity: number;
    id: string;
    className?: string;
    hideStatistics?: boolean;
    accessibilityLabel?: string;
}

export class BarChartComponent extends React.PureComponent<IBarChartComponentProps> {

    render() {
        let items: JSX.Element[] = [];
        for (let i = 0; i < this.props.items.length; i++) {
            let item: IBarChartItem = this.props.items[i];
            let optionCount = item.quantity;
            let percentage: number = Math.round(this.props.totalQuantity != 0 ? (optionCount / this.props.totalQuantity * 100) : 0);
            let percentageString: string = this.props.getBarPercentageString ? this.props.getBarPercentageString(percentage) : percentage + "%";
            let className = "option-container";
            if (this.props.onItemClicked && optionCount > 0) {
                className = className + " pointer-cursor";
            }
            items.push(
                <div role="listitem"
                    aria-label={this.props.items[i].accessibilityLabel}
                    className={className}
                    onClick={() => {
                        if (this.props.onItemClicked) {
                            this.props.onItemClicked(i);
                        }
                    }}
                >
                    <ShimmerContainer lines={1} width={["50%"]} showShimmer={!!this.props.showShimmer}>
                        <Flex gap="gap.small" vAlign="center">
                            <Text aria-hidden={true} title={item.title} content={item.title} size="medium" className={item.titleClassName} truncated />
                            {!item.hideStatistics &&
                                <>
                                    <FlexItem push>
                                        <Text aria-hidden={true} content={optionCount} size="small" weight="bold" />
                                    </FlexItem>
                                    <Text aria-hidden={true} aria-label={percentageString} content={"(" + percentageString + ")"} size="small" />
                                </>}
                        </Flex>
                    </ShimmerContainer>
                    <ShimmerContainer lines={1} showShimmer={!!this.props.showShimmer}>
                        <div className="option-bar">
                            <div className={item.className + " option-percent"} style={{ width: (optionCount / this.props.totalQuantity * 100) + "%" }} />
                        </div>
                    </ShimmerContainer>
                </div>
            );
        }
        return (
            <Flex role="list" aria-label={this.props.accessibilityLabel} column gap="gap.small">
                {items}
            </Flex>
        );
    }
}
