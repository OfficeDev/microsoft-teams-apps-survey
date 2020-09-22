// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { QuestionView, IRatingAnswerProps } from "./QuestionView";
import { StarRatingView } from "./../../RatingView";
import { Localizer } from "../../../utils/Localizer";
/**
 * This component renders for rating(star range) answers
 */
export class StarRatingAnswerView extends React.Component<IRatingAnswerProps> {

    render() {
        const starRatingView: JSX.Element = <StarRatingView
            max={this.props.count}
            disabled={!this.props.editable}
            defaultValue={this.props.response ? this.props.response as number : 0}
            onChange={(value: number) => {
                this.props.updateResponse(value.toString());
            }}
            isPreview={this.props.isPreview}
            className="rating-response-container"
        />;
        return (
            <QuestionView {...this.props}>
                {this.props.editable ?
                    <div aria-label={this.getAccessibilityLabel()}>{starRatingView}</div>
                    : starRatingView}
            </QuestionView>
        );
    }

    private getAccessibilityLabel = () => {
        const accessibilityLabel: string = Localizer.getString("xOfyStarsSelected", this.props.response ? this.props.response as number : 0, this.props.count);
        return accessibilityLabel;
    }

}
