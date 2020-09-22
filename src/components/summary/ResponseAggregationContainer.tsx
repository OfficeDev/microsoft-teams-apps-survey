// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { BarChartComponent, IBarChartItem } from "./../BarChartComponent";
import { Divider, Flex, Text } from "@fluentui/react-northstar";
import { LikeIcon, StarIcon } from "@fluentui/react-icons-northstar";
import { observer } from "mobx-react";
import "./Summary.scss";
import { QuestionDisplayType } from "../Creation/questionContainer/QuestionDisplayType";
import { setCurrentView, setSelectedQuestionDrillDownInfo } from "../../actions/SummaryActions";
import { SummaryPageViewType, QuestionDrillDownInfo } from "../../store/SummaryStore";
import { SurveyUtils } from "../../utils/SurveyUtils";
import { Localizer } from "../../utils/Localizer";
import { UxUtils } from "./../../utils/UxUtils";
import * as actionSDK from "@microsoft/m365-action-sdk";

export interface IResponseAggregationContainerProps {
    questions: actionSDK.ActionDataColumn[];
    responseAggregates: {};
    totalResponsesCount: number;
}

@observer
export class ResponseAggregationContainer extends React.Component<IResponseAggregationContainerProps> {

    render() {
        const maxSingleDigit = 9;
        let questionsSummaryList = [];
        let questions = this.props.questions;
        let responseAggregates = this.props.responseAggregates;
        /*
         Whenever number of questions cross single digit below class is added to align single and double digits
        */
        let className = questions.length > maxSingleDigit ? "question-number-container" : "";
        for (let i = 0; i < questions.length; i++) {
            let titleClassName: string = "question-title";
            if (!questions[i].allowNullValue) {
                titleClassName = titleClassName + " required";
            }
            let questionResultsData;
            switch (questions[i].valueType) {

                case actionSDK.ActionDataColumnValueType.SingleOption:
                case actionSDK.ActionDataColumnValueType.MultiOption:

                    questionResultsData = responseAggregates.hasOwnProperty(questions[i].name) ? JSON.parse(responseAggregates[questions[i].name]) : {};
                    questionsSummaryList.push(
                        <>
                            <Flex gap="gap.smaller">
                                <Flex className={className}>
                                    <Text content={Localizer.getString("QuestionNumber", i + 1)} className="question-number" />
                                </Flex>
                                <Flex column className="overflow-hidden" fill>
                                    <Text content={questions[i].displayName} className={titleClassName} />
                                    {this.getMCQAggregatedView(questionResultsData, questions[i])}
                                </Flex>
                            </Flex>
                            {i != questions.length - 1 && <Divider />}
                        </>);
                    break;

                case actionSDK.ActionDataColumnValueType.Text:
                case actionSDK.ActionDataColumnValueType.Date:
                case actionSDK.ActionDataColumnValueType.DateTime:

                    questionResultsData = responseAggregates.hasOwnProperty(questions[i].name) ? JSON.parse(responseAggregates[questions[i].name]) : [];
                    let responseCount = 0;
                    for (let i = 0; i < questionResultsData.length; i++) {
                        if (!SurveyUtils.isEmptyOrNull(questionResultsData[i])) {
                            responseCount++;
                        }
                    }
                    questionsSummaryList.push(
                        <>
                            <Flex gap="gap.smaller">
                                <Flex className={className}>
                                    <Text content={Localizer.getString("QuestionNumber", i + 1)} className="question-number" />
                                </Flex>
                                <Flex column>
                                    <Text content={questions[i].displayName} className={titleClassName} />
                                    {this.getTextAggregationView(responseCount, questions[i])}
                                </Flex>
                            </Flex>
                            {i != questions.length - 1 && <Divider />}
                        </>);
                    break;

                case actionSDK.ActionDataColumnValueType.Numeric:

                    questionResultsData = responseAggregates.hasOwnProperty(questions[i].name) ? JSON.parse(responseAggregates[questions[i].name]) : {};

                    questionsSummaryList.push(
                        <>
                            <Flex gap="gap.smaller">
                                <Flex className={className}>
                                    <Text content={Localizer.getString("QuestionNumber", i + 1)} className="question-number" />
                                </Flex>
                                <Flex column>
                                    <Text content={questions[i].displayName} className={titleClassName} />
                                    {this.getNumericResponseAggregationView(questionResultsData, questions[i])}
                                </Flex>
                            </Flex>
                            {i != questions.length - 1 && <Divider />}
                        </>);
            }
        }
        return (<div> {questionsSummaryList}</div >);
    }

    private getLikeDislikeSummaryItem(questionResultsData: JSON, question: actionSDK.ActionDataColumn) {
        let likeCount = 0, likePercentage = 0;
        let dislikeCount = 0, dislikePercentage = 0;

        likeCount = questionResultsData[0] || 0;
        dislikeCount = questionResultsData[1] || 0;
        let totalResponsesForQuestion = likeCount + dislikeCount;
        likePercentage = likeCount == 0 ? 0 : Math.round((likeCount * 100) / (likeCount + dislikeCount));
        dislikePercentage = dislikeCount == 0 ? 0 : Math.round((dislikeCount * 100) / (likeCount + dislikeCount));
        let thumbsUpClasses = "reaction";
        let thumbsDownClasses = "reaction";
        if (likeCount >= dislikeCount) {
            thumbsUpClasses = thumbsUpClasses + " yellow-color";
        }
        if (dislikeCount >= likeCount) {
            thumbsDownClasses = thumbsDownClasses + " yellow-color";
        }

        let view = (<Flex padding="padding.medium" column>
            <Text className="stats-indicator" content={totalResponsesForQuestion === 1 ? Localizer.getString("OneResponse")
                : Localizer.getString("TotalResponsesWithCount", totalResponsesForQuestion)} />
            <Flex gap="gap.medium" padding="padding.medium" >
                <Flex.Item size="size.half">
                    <Flex hAlign="center" column gap="gap.small" className={likeCount > 0 ? "rating-drill-down" : ""} onClick={() => this.setDrillDownInfo(likeCount, question, 0, Localizer.getString("ThumbsUpLabel"))} >
                        <LikeIcon className={thumbsUpClasses} outline size="largest"></LikeIcon>
                        <Flex>
                            <Text content={Localizer.getString("ThumbsUpCounter", likeCount, likePercentage)} />
                        </Flex>
                    </Flex>
                </Flex.Item>

                <Flex.Item size="size.half">
                    <Flex hAlign="center" column gap="gap.small" className={dislikeCount > 0 ? "rating-drill-down" : ""} onClick={() => this.setDrillDownInfo(dislikeCount, question, 1, Localizer.getString("ThumbsDownLabel"))}>
                        <LikeIcon rotate={180} className={thumbsDownClasses} outline size="largest"></LikeIcon>
                        <Flex>
                            <Text content={Localizer.getString("ThumbsDownCounter", dislikeCount, dislikePercentage)} />
                        </Flex>
                    </Flex>
                </Flex.Item>
            </Flex></Flex>);
        return view;
    }

    private getMCQAggregatedView(questionResultsData: JSON, question: actionSDK.ActionDataColumn) {
        let customProps = JSON.parse(question.properties);
        let displayType: QuestionDisplayType = (customProps && customProps.hasOwnProperty("dt")) ? customProps["dt"] : QuestionDisplayType.None;
        if (displayType == QuestionDisplayType.LikeDislike) {
            return (this.getLikeDislikeSummaryItem(questionResultsData, question));
        } else {
            let responsesAsBarChartItems: IBarChartItem[] = [];
            let totalResponsesForQuestion: number = 0;
            let average: number = 0;

            for (let j = 0; j < question.options.length; j++) {
                let option = question.options[j];
                let optionCount = questionResultsData[option.name] || 0;
                average = average + (parseInt(option.name)) * optionCount;
                totalResponsesForQuestion = totalResponsesForQuestion + optionCount;
                responsesAsBarChartItems.push({
                    id: option.name,
                    title: option.displayName,
                    quantity: optionCount,
                    className: "loser"
                });
            }

            let item = (
                <BarChartComponent items={responsesAsBarChartItems}
                    getBarPercentageString={(percentage: number) => {
                        return Localizer.getString("BarPercentage", percentage);
                    }}
                    totalQuantity={this.props.totalResponsesCount}
                    onItemClicked={(choiceIndex) => {
                        let optionCount = responsesAsBarChartItems[choiceIndex].quantity;
                        let title = responsesAsBarChartItems[choiceIndex].title;
                        this.setDrillDownInfo(optionCount, question, choiceIndex, title);
                    }}
                />
            );

            if (displayType === QuestionDisplayType.FiveStar ||
                displayType === QuestionDisplayType.TenStar ||
                displayType === QuestionDisplayType.FiveNumber ||
                displayType === QuestionDisplayType.TenNumber) {
                average = totalResponsesForQuestion === 0 ? 0 : average / totalResponsesForQuestion;
                item = (
                    <>
                        <Flex hAlign="start" vAlign="center" className="stats-indicator">
                            <Text content={average.toFixed(1)} />
                            {(displayType === QuestionDisplayType.FiveStar || displayType === QuestionDisplayType.TenStar) && <StarIcon size="small" className="star-icon-average-rating"></StarIcon>}
                            <Text className="average-rating-text" content={Localizer.getString("AverageRating")} />
                        </Flex>
                        <div className="rating-items">
                            {item}
                        </div>
                    </>
                );
            }
            return (<div className="mcq-summary-item">{item}</div>);
        }
    }

    private getTextAggregationView(responseCount: number, question: actionSDK.ActionDataColumn) {
        let className = "mcq-summary-item question-summary-text";
        if (responseCount === 0) {
            return (
                <Text
                    className={className}
                    content={Localizer.getString("ZeroResponses")}
                />);
        }
        className = className + " underline";
        //Passing undefined as we have only first 10 responses in summary
        //and we do not know the exact count of responses
        return this.getQuestionDrillDownInfoView(question, className);
    }

    private getNumericResponseAggregationView(questionResultsData: JSON, question: actionSDK.ActionDataColumn) {
        let sum = questionResultsData.hasOwnProperty("s") ? questionResultsData["s"] : 0;
        let average = questionResultsData.hasOwnProperty("a") ? questionResultsData["a"] : 0;
        let responsesCount = (sum === 0) ? this.props.totalResponsesCount : (Math.round(sum / average));
        const sumString = <Text content={Localizer.getString("Sum", sum)} />;
        const averageString = <Text content={Localizer.getString("Average", average.toFixed(2))} />;
        let className = "";
        if (responsesCount > 0) {
            className = className + " underline";
        }
        if (UxUtils.renderingForMobile()) {
            return (
                <Flex className="stats-indicator mcq-summary-item" column>
                    {this.getQuestionDrillDownInfoView(question, className, responsesCount)}
                    {sumString}
                    {averageString}
                </Flex>
            );
        }
        return (
            <Flex gap="gap.medium" className="stats-indicator mcq-summary-item">
                {this.getQuestionDrillDownInfoView(question, className, responsesCount)}
                <span className="vertical-divider" />
                {sumString}
                <span className="vertical-divider" />
                {averageString}
            </Flex>
        );
    }

    private setDrillDownInfo(responseCount: number, question: actionSDK.ActionDataColumn, choiceIndex: number, subTitle: string) {
        if (responseCount !== 0) {
            let questionDrillDownInfo: QuestionDrillDownInfo = {
                id: parseInt(question.name),
                title: question.displayName,
                type: question.valueType,
                responseCount: responseCount,
                choiceIndex: choiceIndex,
                displayType: JSON.parse(question.properties)["dt"],
                subTitle: subTitle
            };
            setSelectedQuestionDrillDownInfo(questionDrillDownInfo);
            setCurrentView(SummaryPageViewType.ResponseAggregationView);
        }
    }

    private getQuestionDrillDownInfoView(question: actionSDK.ActionDataColumn, className: string, responseCount?: number) {
        const responseCountString = responseCount === undefined ? Localizer.getString("ViewResponses") :
            responseCount === 1 ? Localizer.getString("OneResponse")
                : Localizer.getString("TotalResponsesWithCount", responseCount);
        return (
            <Flex
                onClick={() => {
                    this.setDrillDownInfo(responseCount, question, -1, responseCount === undefined ? undefined : responseCount === 1 ?
                        Localizer.getString("OneResponse") : Localizer.getString("TotalResponsesWithCount", responseCount));
                }}
                className={className} vAlign="center"
                {...(responseCount > 0 && UxUtils.getTabKeyProps())}
                aria-label={responseCountString} >
                <Text weight="regular"
                    className="question-summary-text"
                    content={responseCountString}
                />
            </Flex>
        );
    }

}
