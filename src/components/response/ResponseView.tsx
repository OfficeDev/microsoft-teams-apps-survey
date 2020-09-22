// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ResponseViewMode } from "../../store/ResponseStore";
import { SurveyUtils } from "../../utils/SurveyUtils";
import { QuestionDisplayType } from "../Creation/questionContainer/QuestionDisplayType";
import { IRatingAnswerProps, IQuestionProps, IMultiChoiceProps } from "./questionView/QuestionView";
import { ScaleRatingAnswerView } from "./questionView/ScaleRatingAnswerView";
import { StarRatingAnswerView } from "./questionView/StarRatingAnswerView";
import { Text } from "@fluentui/react-northstar";
import { DateOnlyAnswerView } from "./questionView/DateOnlyAnswerView";
import { NumericAnswerView } from "./questionView/NumericAnswerView";
import { TextAnswerView } from "./questionView/TextAnswerView";
import { MultiSelectView } from "./questionView/MultiSelectView";
import { SingleSelectView } from "./questionView/SingleSelectView";
import { LikeToggleRatingAnswerView } from "./questionView/LikeToggleAnswerView";
import "./Response.scss";
import * as ReactDOM from "react-dom";
import { updateTopMostErrorIndex } from "../../actions/ResponseActions";
import { Localizer } from "../../utils/Localizer";
import { UxUtils } from "./../../utils/UxUtils";
import { Constants } from "./../../utils/Constants";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

export interface IResponseView {
    questionNumber: number;
    actionInstanceColumn: actionSDK.ActionDataColumn;
    response?: any;
    callback?: (response: any) => void;
    responseState: ResponseViewMode;
    isValidationModeOn?: boolean;
    locale?: string;
    setErroredFocus?: boolean;
}

export default class ResponseView extends React.Component<IResponseView, any> {

    shouldComponentUpdate(nextProps: IResponseView) {
        if (this.props.responseState === ResponseViewMode.CreationPreview && nextProps.responseState === ResponseViewMode.CreationPreview) {
            if (this.props.questionNumber == nextProps.questionNumber
                && JSON.stringify(this.props.actionInstanceColumn) == JSON.stringify(nextProps.actionInstanceColumn)) {
                return false;
            }
        }
        return true;
    }

    setErroredFocus() {
        const node = ReactDOM.findDOMNode(this) as HTMLElement;
        UxUtils.setFocus(node, [Constants.FOCUSABLE_ITEMS.INPUT, Constants.FOCUSABLE_ITEMS.TEXTAREA, Constants.FOCUSABLE_ITEMS.TAB]);
        updateTopMostErrorIndex(-1);
    }

    render() {
        ActionSdkHelper.hideLoadIndicator();
        let errorString = this.props.isValidationModeOn
            && !SurveyUtils.isValidResponse(this.props.response, this.props.actionInstanceColumn.allowNullValue, this.props.actionInstanceColumn.valueType) ? Localizer.getString("RequiredAsterisk") : "";
        let questionType: actionSDK.ActionDataColumnValueType = this.props.actionInstanceColumn.valueType;
        if (questionType === actionSDK.ActionDataColumnValueType.Numeric &&
            SurveyUtils.isInvalidNumericPattern(this.props.response)) {
            errorString = Localizer.getString("OnlyNumericAccepted");
        }
        let isPreview = (this.props.responseState === ResponseViewMode.CreationPreview);
        let editable = (this.props.responseState === ResponseViewMode.NewResponse || this.props.responseState === ResponseViewMode.UpdateResponse);
        if (errorString.length > 0 && this.props.setErroredFocus) {
            this.setErroredFocus();
        }

        return (<li className={errorString.length > 0 ? "invalid question-view" : "question-view"} key={this.props.actionInstanceColumn.name} tabIndex={0}>
            {(() => {
                switch (questionType) {
                    case actionSDK.ActionDataColumnValueType.SingleOption: {
                        let displayType: number = JSON.parse(this.props.actionInstanceColumn.properties)["dt"];
                        switch (displayType) {
                            case QuestionDisplayType.TenNumber:
                            case QuestionDisplayType.FiveNumber:
                                let scaleRatingAnswerViewProps: IRatingAnswerProps = {
                                    questionNumber: this.props.questionNumber,
                                    questionText: this.props.actionInstanceColumn.displayName,
                                    editable: editable,
                                    required: !this.props.actionInstanceColumn.allowNullValue,
                                    count: displayType == QuestionDisplayType.TenNumber ? 10 : 5,
                                    response: this.props.response,
                                    updateResponse: this.props.callback,
                                    isPreview: isPreview
                                };
                                return <ScaleRatingAnswerView {...scaleRatingAnswerViewProps} />;

                            case QuestionDisplayType.TenStar:
                            case QuestionDisplayType.FiveStar:
                                let starRatingAnswerViewProps: IRatingAnswerProps = {
                                    questionNumber: this.props.questionNumber,
                                    questionText: this.props.actionInstanceColumn.displayName,
                                    editable: editable,
                                    required: !this.props.actionInstanceColumn.allowNullValue,
                                    count: displayType == QuestionDisplayType.TenStar ? 10 : 5,
                                    response: this.props.response,
                                    updateResponse: this.props.callback,
                                    isPreview: isPreview
                                };
                                return <StarRatingAnswerView {...starRatingAnswerViewProps} />;

                            case QuestionDisplayType.LikeDislike:
                                let likeToggleAnswerViewProps: IQuestionProps = {
                                    questionNumber: this.props.questionNumber,
                                    questionText: this.props.actionInstanceColumn.displayName,
                                    editable: editable,
                                    required: !this.props.actionInstanceColumn.allowNullValue,
                                    response: this.props.response,
                                    updateResponse: this.props.callback,
                                    isPreview: isPreview
                                };
                                return <LikeToggleRatingAnswerView {...likeToggleAnswerViewProps} />;

                            default:
                                let singleSelectProps: IMultiChoiceProps = {
                                    questionNumber: this.props.questionNumber,
                                    questionText: this.props.actionInstanceColumn.displayName,
                                    editable: editable,
                                    required: !this.props.actionInstanceColumn.allowNullValue,
                                    options: this.props.actionInstanceColumn.options,
                                    response: this.props.response,
                                    updateResponse: this.props.callback,
                                    isPreview: isPreview
                                };
                                return <SingleSelectView {...singleSelectProps} />;
                        }

                    }

                    case actionSDK.ActionDataColumnValueType.MultiOption:
                        let multiSelectProps: IMultiChoiceProps = {
                            questionNumber: this.props.questionNumber,
                            questionText: this.props.actionInstanceColumn.displayName,
                            editable: editable,
                            required: !this.props.actionInstanceColumn.allowNullValue,
                            options: this.props.actionInstanceColumn.options,
                            response: this.props.response,
                            updateResponse: this.props.callback,
                            isPreview: isPreview
                        };
                        return <MultiSelectView {...multiSelectProps} />;

                    case actionSDK.ActionDataColumnValueType.Text:
                        let textAnswerProps: IQuestionProps = {
                            questionNumber: this.props.questionNumber,
                            questionText: this.props.actionInstanceColumn.displayName,
                            editable: editable,
                            required: !this.props.actionInstanceColumn.allowNullValue,
                            response: this.props.response,
                            updateResponse: this.props.callback,
                            isPreview: isPreview
                        };
                        return <TextAnswerView {...textAnswerProps} />;

                    case actionSDK.ActionDataColumnValueType.Numeric:
                        let numAnswerProps: IQuestionProps = {
                            questionNumber: this.props.questionNumber,
                            questionText: this.props.actionInstanceColumn.displayName,
                            editable: editable,
                            required: !this.props.actionInstanceColumn.allowNullValue,
                            response: this.props.response,
                            updateResponse: this.props.callback,
                            isPreview: isPreview
                        };
                        return <NumericAnswerView {...numAnswerProps} />;

                    case actionSDK.ActionDataColumnValueType.Date:
                        let dateAnswerProps: IQuestionProps = {
                            questionNumber: this.props.questionNumber,
                            questionText: this.props.actionInstanceColumn.displayName,
                            editable: editable,
                            required: !this.props.actionInstanceColumn.allowNullValue,
                            response: this.props.response,
                            updateResponse: this.props.callback,
                            isPreview: isPreview,
                            locale: this.props.locale
                        };
                        return <DateOnlyAnswerView {...dateAnswerProps} />;

                    default:
                        return null;
                }
            })()}
            {(errorString.length > 0)
                && <Text className="response-mandatory" content={errorString} />
            }
        </li>);
    }
}
