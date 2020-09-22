// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface SurveyViewData {
    // Title
    ti: string;
    // Expiry time
    et: number;
    // Columns(questions)
    cl: string[];
    // rows visibility
    rv: number;
    // Is multi response alloweed
    // mr: number
}

export interface Questions {
    qe: string;
}

export enum SurveyViewDataIndices {
    TitleIndex = 0,
    IsOptionalIndex = 1,
    QuestionDisplayTypeIndex = 2,
    QuestionTypeIndex = 3,
    OptionsIndex = 4
}
