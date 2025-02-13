// function PutFormula() {
//   console.log("Calculating Values, NetRate, and NetValue...");

//   Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     const range = sheet.getUsedRange();
//     range.load("values");
//     await context.sync();

//     const data = range.values;
//     const headers = data[0];

//     //column & headers fields
//     const wgtIndex = headers.indexOf("WGT");
//     const rateIndex = headers.indexOf("RATE");
//     const discPerIndex = headers.indexOf("DISC_PER");
//     const valueIndex = headers.indexOf("VALUE");
//     const netRateIndex = headers.indexOf("NET_RATE");
//     const netValueIndex = headers.indexOf("NET_VALUE");

//     if (
//       wgtIndex === -1 ||
//       rateIndex === -1 ||
//       discPerIndex === -1 ||
//       valueIndex === -1 ||
//       netRateIndex === -1 ||
//       netValueIndex === -1
//     ) {
//       console.error("Required columns not found in the sheet.");
//       return;
//     }

//     for (let i = 1; i < data.length; i++) {
//       const wgt = data[i][wgtIndex];
//       const rate = data[i][rateIndex];
//       const discPer = data[i][discPerIndex];

//       if (wgt !== undefined && rate !== undefined && discPer !== undefined) {
//         const value = wgt * rate;
//         const netRate = rate + (rate * discPer) / 100;
//         const netValue = wgt * netRate;

//         // calculated columns for results
//         data[i][valueIndex] = value;
//         data[i][netRateIndex] = netRate;
//         data[i][netValueIndex] = netValue;
//       }
//     }

//     range.values = data;
//     await context.sync();
//     console.log("Calculations completed.");
//   }).catch("calculated");
// }

// async function PutFormula2() {
//   console.log("Calculating Values, NetRate, and NetValue...");

//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     const range = sheet.getUsedRange();
//     range.load("values");
//     await context.sync();

//     let data = range.values;
//     let headers = data[0].map(header => header.trim().toLowerCase());

//     const synonyms = {
//       WGT: ["total cts", "totalcts", "weight", "wt#", "cts", "carat", "wgt", "size"],
//       RATE: ["baserate", "rate", "raprate", "rap_price", "list rate", "price", "r.price", "rap list"],
//       DISC_PER: ["back", "disc %", "discount", "rap%", "saleback", "off %"],
//     };

//     //Function to find column index dynamically
//     function findColumnIndex(synonymsArray) {
//       return headers.findIndex(header =>
//         synonymsArray.some(synonym => header.includes(synonym))
//       );
//     }

//     const wgtIndex = findColumnIndex(synonyms.WGT);
//     const rateIndex = findColumnIndex(synonyms.RATE);
//     const discPerIndex = findColumnIndex(synonyms.DISC_PER);
//     const valueIndex = headers.indexOf("value");
//     const netRateIndex = headers.indexOf("net_rate");
//     const netValueIndex = headers.indexOf("net_value");

//     if ([wgtIndex, rateIndex, discPerIndex, valueIndex, netRateIndex, netValueIndex].includes(-1)) {
//       console.error("⚠ Required columns not found. Check if synonyms match the Excel headers.");
//       return;
//     }

//     //  Apply calculations row by row
//     for (let i = 1; i < data.length; i++) {
//       let wgt = data[i][wgtIndex];
//       let rate = data[i][rateIndex];
//       let discPer = data[i][discPerIndex];

//       if (wgt !== undefined && rate !== undefined && discPer !== undefined) {
//         // Calculate Values
//         let value = wgt * rate;
//         let netRate = rate + (rate * discPer) / 100;
//         let netValue = wgt * netRate;

//         // Store results in the array
//         data[i][valueIndex] = value;
//         data[i][netRateIndex] = netRate;
//         data[i][netValueIndex] = netValue;
//       }
//     }

//     range.values = data;
//     await context.sync();
//     console.log("✅ Formulas applied successfully!");
//   }).catch(error => console.error("❌ Error in PutFormula:", error));
// }
