// src/backend/handlers.js
import { applySum, applyAverage, sortData } from "./Formulas";

// eslint-disable-next-line no-undef
Office.onReady(() => {
  console.log("Office.js is ready");
});

// Map ribbon button actions
export async function handlePutFormula() {
  // eslint-disable-next-line no-undef
  await Excel.run(async (context) => {
    await applySum(context, "A1:A10"); // Example range
  });
}

export async function handleSort() {
  // eslint-disable-next-line no-undef
  await Excel.run(async (context) => {
    await sortData(context, "A1:D20", 2, "Ascending"); // Example: Sort by the second column
  });
}

export async function handleAverageFormula() {
  await Excel.run(async (context) => {
    await applyAverage(context, "A1:A10");
  });
}
