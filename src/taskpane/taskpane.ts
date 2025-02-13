/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function applySum(context, range) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const targetRange = sheet.getRange(range);
  targetRange.formulas = [["=SUM(A1:A10)"]]; // Replace with dynamic range logic
  await context.sync();
}

export async function applyAverage(context, range) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const targetRange = sheet.getRange(range);
  targetRange.formulas = [["=AVERAGE(A1:A10)"]];
  await context.sync();
}

export async function sortData(context, range, columnIndex, sortOrder = "Ascending") {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const rangeToSort = sheet.getRange(range);
  rangeToSort.sort.apply([{ key: columnIndex, ascending: sortOrder === "Ascending" }]);
  await context.sync();
}
