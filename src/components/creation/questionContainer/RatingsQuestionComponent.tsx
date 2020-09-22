// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "../Creation.scss";
import { Dropdown, Checkbox, Flex, Text, Divider, DropdownItemProps, DropdownProps } from "@fluentui/react-northstar";
import { SurveyUtils } from "../../../utils/SurveyUtils";
import { updateCustomProps } from "../../../actions/CreationActions";
import { QuestionDisplayType } from "./QuestionDisplayType";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { StarRatingView, ToggleRatingView, ScaleRatingView } from "../../RatingView";
import { Localizer } from "../../../utils/Localizer";

export interface IRatingsQuestionComponentProps {
    question: actionSDK.ActionDataColumn;
    questionIndex: number;
    onChange?: (props: IRatingsQuestionComponentProps) => void;
    renderForMobile?: boolean;
}

/**
 * Question component used for rating type question. The format used for rating type is similar to MCQ and
 * options are provided based on the rating level selected
 * Different type of rating questions: star(level/option 5 and 10), scale(level/option 5 and 10) and like/dislike(level/option 2)
 */
export class RatingsQuestionComponent extends React.Component<IRatingsQuestionComponentProps> {

    state = { isDropDownOpen: false };

    typeChoiceSet = [
        {
            index: 0,
            content: Localizer.getString("StarText")
        },
        {
            index: 1,
            content: Localizer.getString("Number")
        },
        {
            index: 2,
            content: Localizer.getString("LikeDislike")
        }
    ];

    levelChoiceSet = [
        {
            content: 5
        },
        {
            content: 10
        }
    ];

    private selectedLevel = JSON.parse(this.props.question.properties)["level"];
    private selectedQuestionType = JSON.parse(this.props.question.properties)["type"];
    private questionDisplayType = JSON.parse(this.props.question.properties)["dt"];

    render() {
        return (
            <div>
                {this.getQuestionView()}
                <Flex column>
                    <Flex className="rating-setting" gap="gap.large">
                        {this.getTypeDropDown()}
                        {this.questionDisplayType !== QuestionDisplayType.LikeDislike ? this.getLevelDropDown() : null}
                    </Flex>
                    <Divider className="question-divider" />
                    {this.getCheckBox()}
                </Flex>
            </div>
        );
    }

    private setQuestionDisplayType() {
        switch (this.selectedQuestionType) {
            case Localizer.getString("StarText"):
                if (this.selectedLevel == 5) {
                    this.questionDisplayType = QuestionDisplayType.FiveStar;
                } else {
                    this.questionDisplayType = QuestionDisplayType.TenStar;
                }
                break;
            case Localizer.getString("Number"):
                if (this.selectedLevel == 5) {
                    this.questionDisplayType = QuestionDisplayType.FiveNumber;
                } else {
                    this.questionDisplayType = QuestionDisplayType.TenNumber;
                }
                break;
            case Localizer.getString("LikeDislike"):
                this.questionDisplayType = QuestionDisplayType.LikeDislike;
        }
    }

    private handleOpenChange = (e, { open }) => {
        this.setState({
            isDropDownOpen: this.selectedQuestionType !== Localizer.getString("LikeDislike") ? open : false,
        });
    }

    private getTypeDropDown = () => {
        let thisProps: IRatingsQuestionComponentProps = {
            question: { ...this.props.question },
            questionIndex: this.props.questionIndex
        };
        let selectedTypeIndex: number = 0;
        let ratingTypes: DropdownItemProps[] = this.typeChoiceSet.map((ratingType, index) => {
            selectedTypeIndex = this.selectedQuestionType === ratingType.content ? index : selectedTypeIndex;
            return {
                header: ratingType.content,
                "aria-label": this.selectedQuestionType === ratingType.content
                    ? Localizer.getString("SelectedRatingType", ratingType.content)
                    : Localizer.getString("UnselectedRatingType", ratingType.content),
                isFromKeyboard: !this.props.renderForMobile //this is set to handle focus of selected item - no changes needed for mobile
            };
        });
        let getA11yStatusMessage = (options) => {
            if (this.props.renderForMobile) {
                return Localizer.getString("DropDownListInfoMobile", ratingTypes.length);
            }
            return Localizer.getString("DropDownListInfo", ratingTypes.length);
        };
        return (
            <Flex gap="gap.medium" className="rating-dropdown-container">
                <Text className="align-center" content={Localizer.getString("TypeText")} />
                <Dropdown value={ratingTypes[selectedTypeIndex]} items={ratingTypes}
                    onChange={(e, props: DropdownProps) => {
                        this.selectedQuestionType = props.value["header"].toString();
                        this.setQuestionDisplayType();
                        let customProps = JSON.parse(thisProps.question.properties);
                        customProps["dt"] = this.questionDisplayType;
                        thisProps.question.properties = JSON.stringify(customProps);
                        thisProps.question.options = SurveyUtils.getRatingQuestionOptions(this.questionDisplayType);
                        this.props.onChange(thisProps);
                        customProps["type"] = this.selectedQuestionType;
                        updateCustomProps(thisProps.questionIndex, customProps);
                    }}
                    getA11yStatusMessage={getA11yStatusMessage}
                    getA11ySelectionMessage={{
                        onAdd: (item: DropdownItemProps) => Localizer.getString("RatingTypeSelected", item.header)
                    }}
                    triggerButton={{ "aria-label": Localizer.getString("RatingType", this.selectedQuestionType) }}
                    inline className="rating-dropdown" />
            </Flex>
        );
    }

    private getLevelDropDown = () => {
        let thisProps: IRatingsQuestionComponentProps = {
            question: { ...this.props.question },
            questionIndex: this.props.questionIndex
        };
        let selectedLevelIndex: number = 0;
        let ratingScales: DropdownItemProps[] = this.levelChoiceSet.map((ratingScale, index) => {
            selectedLevelIndex = this.selectedLevel === ratingScale.content ? index : selectedLevelIndex;
            return {
                header: ratingScale.content.toString(),
                "aria-label": this.selectedLevel === ratingScale.content
                    ? Localizer.getString("SelectedRatingScale", ratingScale.content)
                    : Localizer.getString("UnselectedRatingScale", ratingScale.content),
                isFromKeyboard: !this.props.renderForMobile //this is set to handle focus of selected item - no changes needed for mobile
            };
        });
        let getA11yStatusMessage = (options) => {
            if (this.props.renderForMobile) {
                return Localizer.getString("DropDownListInfoMobile", ratingScales.length);
            }
            return Localizer.getString("DropDownListInfo", ratingScales.length);
        };
        return (
            <Flex gap="gap.medium" className="rating-dropdown-container">
                <Text className="align-center" content={Localizer.getString("ScaleText")} />
                <Dropdown value={ratingScales[selectedLevelIndex]} items={ratingScales}
                    onChange={(e, props: DropdownProps) => {
                        this.selectedLevel = parseInt(props.value["header"]);
                        let customProps = JSON.parse(thisProps.question.properties);
                        this.setQuestionDisplayType();
                        customProps["dt"] = this.questionDisplayType;
                        thisProps.question.properties = JSON.stringify(customProps);
                        thisProps.question.options = SurveyUtils.getRatingQuestionOptions(this.questionDisplayType);
                        this.props.onChange(thisProps);
                        customProps["level"] = this.selectedLevel;
                        updateCustomProps(thisProps.questionIndex, customProps);
                    }}
                    getA11yStatusMessage={getA11yStatusMessage}
                    getA11ySelectionMessage={{
                        onAdd: (item: DropdownItemProps) => Localizer.getString("RatingLevelSelected", item.header)
                    }}
                    triggerButton={{ "aria-label": Localizer.getString("RatingScale", this.selectedLevel) }}
                    inline className="rating-dropdown" open={this.state.isDropDownOpen} onOpenChange={this.handleOpenChange} />
            </Flex>
        );
    }

    private getCheckBox = () => {
        let thisProps: IRatingsQuestionComponentProps = {
            question: { ...this.props.question },
            questionIndex: this.props.questionIndex
        };
        return (
            <Checkbox checked={!(this.props.question.allowNullValue)} label={Localizer.getString("Required")} onChange={(e, data) => {
                thisProps.question.allowNullValue = !(data.checked);
                this.props.onChange(thisProps);
            }} className="required-question-checkbox" />
        );
    }

    private getQuestionView() {
        switch (this.questionDisplayType) {
            case QuestionDisplayType.FiveStar:
            case QuestionDisplayType.TenStar:
                return (<div className="question-preview rating-star"><StarRatingView max={this.selectedLevel} disabled defaultValue={0} isPreview /></div>);
            case QuestionDisplayType.FiveNumber:
            case QuestionDisplayType.TenNumber:
                return (<div className="question-preview"><ScaleRatingView max={this.selectedLevel} disabled defaultValue={0} isPreview /></div>);
            case QuestionDisplayType.LikeDislike:
                return (<div className="question-preview rating-star"><ToggleRatingView isPreview /></div>);
        }
    }
}
