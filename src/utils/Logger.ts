// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Logger component with info and error log levels
 */
export class Logger {

    public static logInfo(message: string) {
        console.info(message);
    }

    public static logError(message: string) {
        console.error(message);
    }
}
