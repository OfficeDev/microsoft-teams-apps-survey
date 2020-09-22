// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { RecyclerListView, DataProvider, LayoutProvider } from "recyclerlistview/web";
import { Flex } from "@fluentui/react-northstar";
import "./RecyclerViewComponent.scss";

export interface IRecyclerViewComponentProps<T> {
    rowHeight: number;
    data: T[];
    showHeader?: React.Key;
    showFooter?: React.Key;
    nonDeterministicRendering?: boolean;
    gridWidth?: number;
    onRowRender: (type: RecyclerViewType, index: number, dataProps: T) => JSX.Element;
}
export enum RecyclerViewType {
    Header,
    Item,
    Footer
}

export class RecyclerViewComponent<T> extends React.Component<IRecyclerViewComponentProps<T>> {
    private layoutProvider: LayoutProvider = null;
    private dataProvider = new DataProvider((r1: T, r2: T) => {
        return r1 !== r2;
    });

    constructor(props: IRecyclerViewComponentProps<T>) {
        super(props);
        this.initialize(props);
    }

    shouldComponentUpdate(nextProps: IRecyclerViewComponentProps<T>) {
        if (nextProps !== this.props) {
            this.updateDataProvider(nextProps);
        }
        return true;
    }

    render() {
        return (
            <Flex fill column className="recycler-container">
                <RecyclerListView
                    key={this.props.gridWidth}
                    rowHeight={this.props.rowHeight}
                    layoutProvider={this.layoutProvider}
                    dataProvider={this.dataProvider}
                    rowRenderer={(type: RecyclerViewType, data: T, index: number): JSX.Element => {
                        return this.props.onRowRender(type, index, data);
                    }}
                    {...this.props.nonDeterministicRendering ? this.getViewProps() : {}}
                />
            </Flex>);
    }

    private initialize(props: IRecyclerViewComponentProps<T>) {
        //Create the layout provider
        //First method: Given an index return the type of item e.g ListItemType1, ListItemType2 in case you have variety of items in your list/grid
        //Second: Given a type and object set the height and width for that type on given object
        this.layoutProvider = new LayoutProvider(
            (index: number) => {
                if (this.props.showHeader && index == 0) {
                    return RecyclerViewType.Header;
                } else if (this.props.showFooter && index == this.dataProvider.getSize() - 1) {
                    return RecyclerViewType.Footer;
                } else {
                    return RecyclerViewType.Item;
                }
            },
            (type: number, dim: any) => {
                if (this.props.gridWidth) {
                    dim.width = this.props.gridWidth;
                } else {
                    dim.width = window.innerWidth;
                }
                dim.height = this.props.rowHeight;
            }
        );
        this.updateDataProvider(props);
    }

    private updateDataProvider(props: IRecyclerViewComponentProps<T>) {
        let data: T[] = props.data;
        let dataRow: any[] = [];
        if (props.showHeader) {
            dataRow.push(props.showHeader);
        }
        dataRow = dataRow.concat(data);

        if (props.showFooter) {
            dataRow.push(props.showFooter);
        }
        this.dataProvider = this.dataProvider.cloneWithRows(dataRow);
    }

    private getViewProps() {
        return {
            forceNonDeterministicRendering: true
        };
    }
}
