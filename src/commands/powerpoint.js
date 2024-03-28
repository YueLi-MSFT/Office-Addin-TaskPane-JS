/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    // Register the function with Office.
    Office.actions.associate("action", actionPowerPoint);
  }
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function actionPowerPoint(event) {
    /**
   * Insert your PowerPoint code here
   */
    const options = { coercionType: Office.CoercionType.Text };

    await Office.context.document.setSelectedDataAsync(" ", options);
    await Office.context.document.setSelectedDataAsync("Hello World!", options);
}