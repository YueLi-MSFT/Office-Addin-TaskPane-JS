/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Register the function with Office.
    Office.actions.associate("action", actionExcel);
  }
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function actionExcel(event) {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
