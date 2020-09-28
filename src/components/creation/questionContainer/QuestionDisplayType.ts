// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Rating questions are represented with the similar datastructure as the MCQ single select, so it is used
 * to differentiate between the MCQ single select and rating question types
 */
export enum QuestionDisplayType {
    // Default type
    None = -1,
    Select = 0,
    FiveStar = 1,
    TenStar = 2,
    LikeDislike = 3,
    FiveNumber = 4,
    TenNumber = 5
}
