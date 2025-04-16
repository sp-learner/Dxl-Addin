/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

// async function PutFormula() {
//   console.log("Applying fully dynamic formulas...");

//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     const range = sheet.getUsedRange();
//     range.load("values");
//     await context.sync();

//     let data = range.values;
//     if (data.length === 0 || data[0].length === 0) {
//       console.error("No data found in the sheet.");
//       return;
//     }

//     // Header Row Function
//     let headerRowIndex = -1;
//     for (let i = 0; i < data.length; i++) {
//       if (data[i].some(cell => typeof cell === "string" && cell.trim() !== "")) {
//         headerRowIndex = i;
//         break;
//       }
//     }

//     if (headerRowIndex === -1) {
//       console.error("No valid header row found.");
//       return;
//     }

//     let headers = data[headerRowIndex].map(header => header ? header.toString().trim().toLowerCase() : "");

//     //Synonyms
//     const Synonyms = {
//       WGT: ["TOTAL CTS","TotalCts", "Weight R","weigh", "Cts#", "SIZE#","Wt#", "Car", "Cara", "Carat", "CARATS", "Crt", "Crts", "CRTWT", "CT", "Ct.", "Cts", "Cts.", "POLISE" ,"CT" ,"Size", "SIZE." ,"Weight", "Weight ??", "Wgt" ,"WHT.", "WT", "Wt."],
//       RATE: ["BaseRate", "Disc Price"," Full Rap Price", "List", "List Price", "List Price ????", "List Rate", "LiveRAP", "NEW RAP", "Orap", "price", "R.PRICE", "Rap", "Rap $", "Rap $/CT", "Rap List", "Rap Price", "Rap Price($)", "Rap Rate", "RAP RTE", "Rap$", "RAP($)", "Rap-Price", "RAP.", "Rap.", "Price", "Rap.($)", "Rap/Price", "Rap_per_Crt", "RAP_PRICE", "Rapa", "Rapa Rate", "Rapa_Rate", "rapaport", "RAPAPORT_RATE", "RapaportPrice", "RapaRate", "RapDown", "Rape", "RapList", "RapNet Price", "rapnetcaratprice", "RapNetPrice", "RAPO", "RAPPLIST", "rapprice", "RapRat", "RapRate", "RapRice", "RapRte", "Rate", "repRate"],
//       DISC_PER: ["%"," % Back"," % BELOW", "%Rap", "Asking Disc. %", "Back", "BACK %", "Back (-%)", "Back %", "Back -%", "Back%", "Base Off %", "Base Off%", "CBack", "DIC.", "DIS", "Dis %", "Dis%", "DIS.", "Disc", "Disc %", "Disc%", "Disc(%)", "DISC.", "Disc/Pre", "DISC_PER", "Disco%", "DISCOUNT", "Discount %","Discount % ??", "Discount%", "Discprct", "F disc", "Fair/Last Bid %", "Final %", "Final Disc%", "final_discount", "ListDisc%", "Net %", "New Rap%", "Off %", "Off%", "Offer Disc.(%)", "OffPer", "Price", "R.Dn", "Rap %", "RAP DIS", "Rap Disc", "Rap Disc %", "Rap Discount", "Rap%", "Rap.%", "RAP_DISCOUNT", "rap_per", "RapDis", "RapDown", "rapnet", "Rapnet", "Discount %", "RapNet Back", "Rapnet Discount", "Rapnet Discount%", "rapnetdiscount", "RapnetDiscountPercent", "RapOff", "RP Disc", "saleback", "SaleDis", "SaleDisc", "Selling Disc", "User Disc", "VDisc %"," WebsiteDiscount", "Rapdisc"],
//       Value: ["value", "rapvalue", "rapaport value", "r.value", "val", "RapVlu"],
//     };

//     function findColumnIndex(synonymsArray) {
//       return headers.findIndex((header) => synonymsArray.some((synonym) => header.includes(synonym.toLowerCase())));
//     }

//     const wgtIndex = findColumnIndex(Synonyms.WGT);
//     const rateIndex = findColumnIndex(Synonyms.RATE);
//     const discPerIndex = findColumnIndex(Synonyms.DISC_PER);
//     const valueIndex = findColumnIndex(Synonyms.Value);
//     const netRateIndex = headers.indexOf("net_rate");
//     const netValueIndex = headers.indexOf("net_value");

//     if ([wgtIndex, rateIndex, discPerIndex, valueIndex, netRateIndex, netValueIndex].includes(-1)) {
//       console.error("Required columns not found.");
//       return;
//     }

//     //for finding Column & Row
//     const wgt = getColumnLetter(wgtIndex);
//     const rate = getColumnLetter(rateIndex);
//     const discper = getColumnLetter(discPerIndex);
//     const netRate = getColumnLetter(netRateIndex);

//     for (let i = headerRowIndex + 1; i < data.length; i++) {
//       const rowNum = i + 1;

//       //formulas & Calculations
//       const valueFormula = `=${wgt}${rowNum}*${rate}${rowNum}`;
//       const netRateFormula = `=${rate}${rowNum}+((${rate}${rowNum}*${discper}${rowNum})/100)`;
//       const netValueFormula = `=${wgt}${rowNum}*${netRate}${rowNum}`;

//       data[i][valueIndex] = valueFormula;
//       data[i][netRateIndex] = netRateFormula;
//       data[i][netValueIndex] = netValueFormula;
//     }

//     range.formulas = data;
//     await context.sync();
//     console.log("Fully dynamic formulas applied!");
//   }).catch(error => console.error("Error in PutFormula:", error));
// }

// function getColumnLetter(index) {
//   if (index < 0) return "";
//   let columnLetter = "";
//   let tempIndex = index;

//   while (tempIndex >= 0) {
//     columnLetter = String.fromCharCode((tempIndex % 26) + 65) + columnLetter;
//     tempIndex = Math.floor(tempIndex / 26) - 1;
//   }

//   return columnLetter;
// }

// window.PutFormula = PutFormula;
