/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */

// async function PutFormula() {
//   console.log("Calculating Values, NetRate, and NetValue...");

//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     const range = sheet.getUsedRange();
//     range.load("values");
//     await context.sync();

//     let data = range.values;
//     let headers = data[0].map((header) => header.trim().toLowerCase());

//     //Synonyms
//     const Synonyms = {
//       WGT: ["TotalCts", "total cts", "totalcts", "weight r", "weigh", "cts#", "size#", "wt#", "car", "cara", "carat", "carats", "crt", "crts", "crtwt", "ct", "ct.", "cts", "cts.", "polise ct", "size", "size.", "weight", "weight ??", "wgt", "wht.", "wt", "wt."],
//       RATE: ["raprate", "baserate", "disc price", "full rap price", "list", "list price", "list price ????", "list rate", "liverap", "new rap", "orap", "price", "r.price", "rap", "rap $", "rap $/ct", "rap list", "rap price", "rap price($)", "rap rate", "rap rte", "rap$", "rap($)", "rap-price", "rap.", "rap.", "price", "rap.($)", "rap/price", "rap_per_crt", "rap_price", "rapa", "rapa rate", "rapa_rate", "rapaport", "rapaport_rate", "rapaportprice", "raparate", "rapdown", "rape", "raplist", "rapnet price", "rapnet price ($)", "rapnetcaratprice", "rapnetprice", "rapo", "rapplist", "rapprice", "raprat", "raprate", "raprice", "raprte", "rate", "reprate"],
//       DISC_PER: ["%", "% back", "% below", "%rap", "asking disc. %", "back", "back %", "back (-%)", "back %", "back -%", "back%", "base off %", "base off%", "cback", "dic.", "dis", "dis %", "dis%", "dis.", "disc", "disc %", "disc%", "disc(%)", "disc.", "disc/pre", "disc_per", "disco%", "discount", "discount %", "discount % ??", "discount%", "discprct", "f disc", "fair/last bid %", "final %", "final disc%", "final_discount", "listdisc%", "net %", "new rap%", "off %", "off%", "offer disc.(%)", "offper", "price", "r.dn", "rap %", "rap dis", "rap disc", "rap disc %", "rap discount", "rap%", "rap.%", "rap_discount", "rap_per", "rapdis", "rapdown", "rapnet", "rapnet discount %", "rapnet back", "rapnet discount", "rapnet discount%", "rapnetdiscount", "rapnetdiscountpercent", "rapoff", "rp disc", "saleback", "saledis", "saledisc", "selling disc", "user disc", "vdisc %", "websitediscount", "rapdisc"]
//     };

//     function findColumnIndex(synonymsArray) {
//       return headers.findIndex((header) => synonymsArray.some((synonym) => header.includes(synonym)));
//     }

//     const wgtIndex = findColumnIndex(Synonyms.WGT);
//     const rateIndex = findColumnIndex(Synonyms.RATE);
//     const discPerIndex = findColumnIndex(Synonyms.DISC_PER);
//     const valueIndex = headers.indexOf("value");
//     const netRateIndex = headers.indexOf("net_rate");
//     const netValueIndex = headers.indexOf("net_value");

//     if ([wgtIndex, rateIndex, discPerIndex, valueIndex, netRateIndex, netValueIndex].includes(-1)) {
//       console.error("Required columns not found.");
//       return;
//     }

//     for (let i = 1; i < data.length; i++) {
//       let wgt = parseFloat(data[i][wgtIndex]) || 0;
//       let rate = parseFloat(data[i][rateIndex]) || 0;
//       let discPer = parseFloat(data[i][discPerIndex]) || 0;

//       if (!isNaN(wgt) && !isNaN(rate) && !isNaN(discPer)) {
//         // Calculate Values
//         let value = wgt * rate;
//         let netRate = rate + (rate * discPer) / 100;
//         let netValue = wgt * netRate;

//         data[i][valueIndex] = value.toFixed(2);
//         data[i][netRateIndex] = netRate.toFixed(2);
//         data[i][netValueIndex] = netValue.toFixed(2);
//       }
//     }

//     range.values = data;
//     await context.sync();
//   }).catch((error) => console.error("Error in PutFormula:", error));
// }

// window.PutFormula = PutFormula;

async function PutAverage() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("rowCount, columnCount, values");
      await context.sync();

      const lastDataRow = range.rowCount;
      const summaryRow = lastDataRow + 2; // Leave one empty row below data

      // Define synonyms (your provided lists)
      const InputSynonyms = {
        WGT: ["Weight", "TOTAL CTS","TotalCts", "Weight R","weigh", "Cts#", "SIZE#","Wt#", "Car", "Cara", "Carat", "CARATS", "Crt", "Crts", "CRTWT", "CT", "Ct.", "Cts", "Cts.", "POLISE" ,"CT" ,"Size", "SIZE." ,"Weight", "Weight ??", "Wgt" ,"WHT.", "WT", "Wt."],
        RATE: ["Rate", "BaseRate", "Disc Price"," Full Rap Price", "List", "List Price", "List Price ????", "List Rate", "LiveRAP", "NEW RAP", "Orap", "price", "R.PRICE", "Rap", "Rap $", "Rap $/CT", "Rap List", "Rap Price", "Rap Price($)", "Rap Rate", "RAP RTE", "Rap$", "RAP($)", "Rap-Price", "RAP.", "Rap.", "Price", "Rap.($)", "Rap/Price", "Rap_per_Crt", "RAP_PRICE", "Rapa", "Rapa Rate", "Rapa_Rate", "rapaport", "RAPAPORT_RATE", "RapaportPrice", "RapaRate", "RapDown", "Rape", "RapList", "RapNet Price", "rapnetcaratprice", "RapNetPrice", "RAPO", "RAPPLIST", "rapprice", "RapRat", "RapRate", "RapRice", "RapRte", "Rate", "repRate"],
        DISC: ["Disc", "%"," % Back"," % BELOW", "%Rap", "Asking Disc. %", "Back", "BACK %", "Back (-%)", "Back %", "Back -%", "Back%", "Base Off %", "Base Off%", "CBack", "DIC.", "DIS", "Dis %", "Dis%", "DIS.", "Disc", "Disc %", "Disc%", "Disc(%)", "DISC.", "Disc/Pre", "DISC_PER", "Disco%", "DISCOUNT", "Discount %","Discount % ??", "Discount%", "Discprct", "F disc", "Fair/Last Bid %", "Final %", "Final Disc%", "final_discount", "ListDisc%", "Net %", "New Rap%", "Off %", "Off%", "Offer Disc.(%)", "OffPer", "Price", "R.Dn", "Rap %", "RAP DIS", "Rap Disc", "Rap Disc %", "Rap Discount", "Rap%", "Rap.%", "RAP_DISCOUNT", "rap_per", "RapDis", "RapDown", "rapnet", "Rapnet", "Discount %", "RapNet Back", "Rapnet Discount", "Rapnet Discount%", "rapnetdiscount", "RapnetDiscountPercent", "RapOff", "RP Disc", "saleback", "SaleDis", "SaleDisc", "Selling Disc", "User Disc", "VDisc %"," WebsiteDiscount", "Rapdisc"],
      };

      const ResultSynonyms = {
        VALUE: ["value", "rapvalue", "rapaport value", "r.value", "val", "RapVlu"],
        NET_RATE: ["Net_Rate", "$ / Carat", "$/Carat", "$/CT", "$/CTS", "$/PC", "Asking Price", "askprice", "BACK P/Ct", "Base Rate", "Cash Price", "CashPrice", "CRate", "Ct/Price", "D.RAP PRICE", "DIS / CT", "Final Rate", "List$/Ct", "Net Rate", "NET_RATE", "P.CARAT", "P/CT", "P/CTS", "Per Crt $", "Per ct", "Per Ct $", "PerCarat", "PerCrt", "PerCts", "PPC", "PPC$", "Pr/Ct", "PRAP($)","PRI/CRT", "Price p.c", "Price $/cts", "Price / Crts", "Price Per Carat", "Price Per Crt", "Price Per Ct", "Price/Carat", "Price/Crt", "Price/Ct", "Price/Ct ($)", "Price/ct.", "Price/Cts", "Price/CTS $", "Price/Cts USD", "Price/Cts.", "PRICE_DOLLAR", "PRICE_PER_CARAT", "Price_Per_Crt", "PricePerCarat", "Rap @", "rap_prc", "RapNet Price", "RapNet Rate", "RATE", "Rate $/CT", "Rate / CT", "Rate ?", "Rate per carat as per Rapnet", "Rate($)", "RATE($/CT)", "Rate/Ct", "RP Price", "RTE", "SaleRate", "sales_price", "Selling Price", "User Price /Cts", "VALLUE", "WebsiteRate"],
        NET_VALUE: ["net_value", "Net_Value", "$ Total", "amont", "AMOUNT", "Amount $", "Amount ?", "Amount US$", "Amount($)", "Amt", "Amt $", "Amt.", "askamount", "Asking Amount", "Back Total", "Base Amt", "CAmount", "DiscountPrice", "EST AMT", "F value", "F.Amt", "FINAL", "Final Amount", "Final Amt", "Final Amt IN $", "Final Price", "Final Value", "FINAL$", "final_amount", "FinalValue", "mspTotal", "Net", "NET VALLUE", "NET $", "Net Amt", "Net Amt($)", "Net Value", "NET_VALUE", "NetAmt", "Offer Value($)", "Rap US $", "Rapa Value", "RapNet Amount", "RapNet Price", "RP Tot$", "SaleAmt", "saledollorprice", "Stone Price", "Stone($)", "T AMT", "T Price", "T VALUE", "T. AMOUNT", "T.Amt", "Tot. Value", "Total", "TOTAL $", "Total $ as per Rapnet", "Total ($)", "TOTAL AMOUNT", "Total Amt", "Total Amt.", "Total Price", "Total$", "total_price", "TotalAmount", "TotalPrice", "TotalValue $", "User Total $", "VALUE_DOLLAR", "WebsiteAmount"],
      };

      // Get column letters for key columns using synonyms
      const columns = {};
      for (let col = 0; col < range.columnCount; col++) {
        const header = (range.values[0][col] || "").toString().trim().toLowerCase();
        const letter = String.fromCharCode(65 + col);

        // Check weight synonyms
        if (InputSynonyms.WGT.some(syn => header.includes(syn.toLowerCase()))) {
          columns.weight = letter;
        }
        else if (InputSynonyms.DISC.some(syn => header.includes(syn.toLowerCase()))) {
          columns.value = letter;
        }
        // Check value synonyms
        else if (ResultSynonyms.VALUE.some(syn => header.includes(syn.toLowerCase()))) {
          columns.value = letter;
        }
        // Check net_value synonyms
        else if (ResultSynonyms.NET_VALUE.some(syn => header.includes(syn.toLowerCase()))) {
          columns.net_value = letter;
        }
        // Check rate synonyms
        else if (InputSynonyms.RATE.some(syn => header.includes(syn.toLowerCase()))) {
          columns.rate = letter;
        }
        // Check net_rate synonyms
        else if (ResultSynonyms.NET_RATE.some(syn => header.includes(syn.toLowerCase()))) {
          columns.net_rate = letter;
        }
      }

      // Apply formulas and highlight results
      for (let col = 0; col < range.columnCount; col++) {
        const letter = String.fromCharCode(65 + col);
        const header = (range.values[0][col] || "").toString().trim().toLowerCase();
        const cell = sheet.getRange(`${letter}${summaryRow + 1}`);

        // Check if current column matches any synonyms
        const isWeight = InputSynonyms.WGT.some(syn => header.includes(syn.toLowerCase()));
        const isDisc = InputSynonyms.DISC.some(syn => header.includes(syn.toLowerCase()));
        const isValue = ResultSynonyms.VALUE.some(syn => header.includes(syn.toLowerCase()));
        const isNetValue = ResultSynonyms.NET_VALUE.some(syn => header.includes(syn.toLowerCase()));
        const isRate = InputSynonyms.RATE.some(syn => header.includes(syn.toLowerCase()));
        const isNetRate = ResultSynonyms.NET_RATE.some((syn) => header.includes(syn.toLowerCase()));

        if (isWeight || isValue || isNetValue) {
          // SUM formula for weight/value/net_value columns
          cell.formulas = [[`=SUM(${letter}2:${letter}${lastDataRow})`]];
          cell.format.fill.color = "yellow";
        } 
        else if (isDisc) {
          cell.formulas = [[`=AVERAGE(${letter}2:${letter}${lastDataRow})`]];
          cell.format.fill.color = "yellow";
        }
        else if (isRate && columns.value && columns.weight) {
          // value/weight for rate column
          cell.formulas = [[`=${columns.value}${summaryRow + 1}/${columns.weight}${summaryRow + 1}`]];
          cell.format.fill.color = "yellow";
        }
        else if (isNetRate && columns.net_value && columns.weight) {
          // net_value/weight for net_rate column
          cell.formulas = [[`=${columns.net_value}${summaryRow + 1}/${columns.weight}${summaryRow + 1}`]];
          cell.format.fill.color = "yellow";
        }
      }

      await context.sync();
      showToastNotification("Averages calculated successfully!");
    });
  } catch (error) {
    console.error("Error:", error);
    showToastNotification("Failed to calculate averages", "error");
  }
}

// Helper function to detect column type from synonyms
function detectColumnType(header) {
  const headerLower = header.toLowerCase();

  const InputSynonyms = {
    WGT: ["Weight", "TOTAL CTS","TotalCts", "Weight R","weigh", "Cts#", "SIZE#","Wt#", "Car", "Cara", "Carat", "CARATS", "Crt", "Crts", "CRTWT", "CT", "Ct.", "Cts", "Cts.", "POLISE" ,"CT" ,"Size", "SIZE." ,"Weight", "Weight ??", "Wgt" ,"WHT.", "WT", "Wt."],
    RATE: ["Rate", "BaseRate", "Disc Price"," Full Rap Price", "List", "List Price", "List Price ????", "List Rate", "LiveRAP", "NEW RAP", "Orap", "price", "R.PRICE", "Rap", "Rap $", "Rap $/CT", "Rap List", "Rap Price", "Rap Price($)", "Rap Rate", "RAP RTE", "Rap$", "RAP($)", "Rap-Price", "RAP.", "Rap.", "Price", "Rap.($)", "Rap/Price", "Rap_per_Crt", "RAP_PRICE", "Rapa", "Rapa Rate", "Rapa_Rate", "rapaport", "RAPAPORT_RATE", "RapaportPrice", "RapaRate", "RapDown", "Rape", "RapList", "RapNet Price", "rapnetcaratprice", "RapNetPrice", "RAPO", "RAPPLIST", "rapprice", "RapRat", "RapRate", "RapRice", "RapRte", "Rate", "repRate"],
    DISC: ["Disc", "%"," % Back"," % BELOW", "%Rap", "Asking Disc. %", "Back", "BACK %", "Back (-%)", "Back %", "Back -%", "Back%", "Base Off %", "Base Off%", "CBack", "DIC.", "DIS", "Dis %", "Dis%", "DIS.", "Disc", "Disc %", "Disc%", "Disc(%)", "DISC.", "Disc/Pre", "DISC_PER", "Disco%", "DISCOUNT", "Discount %","Discount % ??", "Discount%", "Discprct", "F disc", "Fair/Last Bid %", "Final %", "Final Disc%", "final_discount", "ListDisc%", "Net %", "New Rap%", "Off %", "Off%", "Offer Disc.(%)", "OffPer", "Price", "R.Dn", "Rap %", "RAP DIS", "Rap Disc", "Rap Disc %", "Rap Discount", "Rap%", "Rap.%", "RAP_DISCOUNT", "rap_per", "RapDis", "RapDown", "rapnet", "Rapnet", "Discount %", "RapNet Back", "Rapnet Discount", "Rapnet Discount%", "rapnetdiscount", "RapnetDiscountPercent", "RapOff", "RP Disc", "saleback", "SaleDis", "SaleDisc", "Selling Disc", "User Disc", "VDisc %"," WebsiteDiscount", "Rapdisc"],
  };

  const ResultSynonyms = {
    VALUE: ["value", "rapvalue", "rapaport value", "r.value", "val", "RapVlu"],
    NET_RATE: ["Net_Rate", "$ / Carat", "$/Carat", "$/CT", "$/CTS", "$/PC", "Asking Price", "askprice", "BACK P/Ct", "Base Rate", "Cash Price", "CashPrice", "CRate", "Ct/Price", "D.RAP PRICE", "DIS / CT", "Final Rate", "List$/Ct", "Net Rate", "NET_RATE", "P.CARAT", "P/CT", "P/CTS", "Per Crt $", "Per ct", "Per Ct $", "PerCarat", "PerCrt", "PerCts", "PPC", "PPC$", "Pr/Ct", "PRAP($)","PRI/CRT", "Price p.c", "Price $/cts", "Price / Crts", "Price Per Carat", "Price Per Crt", "Price Per Ct", "Price/Carat", "Price/Crt", "Price/Ct", "Price/Ct ($)", "Price/ct.", "Price/Cts", "Price/CTS $", "Price/Cts USD", "Price/Cts.", "PRICE_DOLLAR", "PRICE_PER_CARAT", "Price_Per_Crt", "PricePerCarat", "Rap @", "rap_prc", "RapNet Price", "RapNet Rate", "RATE", "Rate $/CT", "Rate / CT", "Rate ?", "Rate per carat as per Rapnet", "Rate($)", "RATE($/CT)", "Rate/Ct", "RP Price", "RTE", "SaleRate", "sales_price", "Selling Price", "User Price /Cts", "VALLUE", "WebsiteRate"],
    NET_VALUE: ["net_value", "Net_Value", "$ Total", "amont", "AMOUNT", "Amount $", "Amount ?", "Amount US$", "Amount($)", "Amt", "Amt $", "Amt.", "askamount", "Asking Amount", "Back Total", "Base Amt", "CAmount", "DiscountPrice", "EST AMT", "F value", "F.Amt", "FINAL", "Final Amount", "Final Amt", "Final Amt IN $", "Final Price", "Final Value", "FINAL$", "final_amount", "FinalValue", "mspTotal", "Net", "NET VALLUE", "NET $", "Net Amt", "Net Amt($)", "Net Value", "NET_VALUE", "NetAmt", "Offer Value($)", "Rap US $", "Rapa Value", "RapNet Amount", "RapNet Price", "RP Tot$", "SaleAmt", "saledollorprice", "Stone Price", "Stone($)", "T AMT", "T Price", "T VALUE", "T. AMOUNT", "T.Amt", "Tot. Value", "Total", "TOTAL $", "Total $ as per Rapnet", "Total ($)", "TOTAL AMOUNT", "Total Amt", "Total Amt.", "Total Price", "Total$", "total_price", "TotalAmount", "TotalPrice", "TotalValue $", "User Total $", "VALUE_DOLLAR", "WebsiteAmount"],
  };
  
  // Check InputSynonyms
  if (InputSynonyms.WGT.some(syn => headerLower.includes(syn.toLowerCase()))) return 'weight';
  if (InputSynonyms.RATE.some(syn => headerLower.includes(syn.toLowerCase()))) return 'rate';
  if (InputSynonyms.DISC.some(syn => headerLower.includes(syn.toLowerCase()))) return 'disc';
  
  // Check ResultSynonyms
  if (ResultSynonyms.VALUE.some(syn => headerLower.includes(syn.toLowerCase()))) return 'value';
  if (ResultSynonyms.NET_RATE.some(syn => headerLower.includes(syn.toLowerCase()))) return 'net_rate';
  if (ResultSynonyms.NET_VALUE.some(syn => headerLower.includes(syn.toLowerCase()))) return 'net_value';
  
  return 'other';
}

// Helper function to check if column contains numeric data
async function isNumericColumn(sheet, columnIndex) {
  const testRange = sheet.getRangeByIndexes(1, columnIndex, 5, 1);
  testRange.load("values");
  await testRange.context.sync();
  return testRange.values.some(row => !isNaN(parseFloat(row[0])));
}

// function handleAvarageFormula(){
//     Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const range = sheet.getUsedRange();
//         range.load("rowCount, columnCount, values");
//         await context.sync();

//         const rowCount = range.rowCount;

//         for(let i =1; i < rowCount; i++){
//             const valueCell = sheet.getCell(i,3);
//             const netRateCell = sheet.getCell(i, 4);
//             const netValueCell = sheet.getCell(i, 5);

//             sheet.getCell(rowCount, 0).formulas = [
//                 [`=Avarage(A2:A${rowCount})`] // Sum of WGT
//             ];
//         }
//     })
// }

// function handleCompleteCalculations() {
//     console.log("Performing complete calculations...");

//     Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const range = sheet.getUsedRange();
//         range.load("values, rowCount, columnCount");
//         await context.sync();

//         const data = range.values;
//         const headers = data[0]; // Assuming headers are in the first row

//         // Find the indices of required columns
//         const wgtIndex = headers.indexOf("WGT");
//         const valueIndex = headers.indexOf("VALUE");
//         const netValueIndex = headers.indexOf("NET_VALUE");
//         const rateIndex = headers.indexOf("RATE");
//         const netRateIndex = headers.indexOf("NET_RATE");

//         if (
//             wgtIndex === -1 ||
//             valueIndex === -1 ||
//             netValueIndex === -1 ||
//             rateIndex === -1 ||
//             netRateIndex === -1
//         ) {
//             console.error("Required columns not found. Please ensure correct headers.");
//             return;
//         }

//         // Row for inserting calculations (next row after data)
//         const calculationRow = data.length; // Last row + 1 (0-based index)

//         // Write formulas for SUM in WGT, VALUE, and NET_VALUE columns
//         sheet.getCell(calculationRow, wgtIndex).formulas = [
//             `=SUM(${sheet.getRangeByIndexes(1, wgtIndex, data.length - 1, 1).getAddress()})`
//         ];
//         sheet.getCell(calculationRow, valueIndex).formulas = [
//             `=SUM(${sheet.getRangeByIndexes(1, valueIndex, data.length - 1, 1).getAddress()})`
//         ];
//         sheet.getCell(calculationRow, netValueIndex).formulas = [
//             `=SUM(${sheet.getRangeByIndexes(1, netValueIndex, data.length - 1, 1).getAddress()})`
//         ];

//         // Write division formulas in RATE and NET_RATE columns
//         sheet.getCell(calculationRow, rateIndex).formulas = [
//             `=${sheet.getCell(calculationRow, wgtIndex).getAddress()} / ${sheet.getCell(
//                 calculationRow,
//                 valueIndex
//             ).getAddress()}`
//         ];
//         sheet.getCell(calculationRow, netRateIndex).formulas = [
//             `=${sheet.getCell(calculationRow, wgtIndex).getAddress()} / ${sheet.getCell(
//                 calculationRow,
//                 netValueIndex
//             ).getAddress()}`
//         ];

//         await context.sync();
//         console.log("Calculations and formulas applied successfully.");
//     }).catch((error) => {
//         console.error(error);
//     });
// }

async function PutFormula() {
  console.log("Applying formulas...");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    let data = range.values;
    if (data.length === 0 || data[0].length === 0) {
      console.error("No data found.");
      return;
    }

    // Find header row (first non-empty row)
    let headerRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(cell => typeof cell === "string" && cell.trim() !== "")) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      console.error("No header row found.");
      return;
    }

    let headers = data[headerRowIndex].map(header => header ? header.toString().trim().toLowerCase() : "");

    const InputSynonyms = {
      WGT: ["Weight", "TOTAL CTS","TotalCts", "Weight R","weigh", "Cts#", "SIZE#","Wt#", "Car", "Cara", "Carat", "CARATS", "Crt", "Crts", "CRTWT", "CT", "Ct.", "Cts", "Cts.", "POLISE" ,"CT" ,"Size", "SIZE." ,"Weight", "Weight ??", "Wgt" ,"WHT.", "WT", "Wt."],
      RATE: ["Rate", "BaseRate", "Disc Price"," Full Rap Price", "List", "List Price", "List Price ????", "List Rate", "LiveRAP", "NEW RAP", "Orap", "price", "R.PRICE", "Rap", "Rap $", "Rap $/CT", "Rap List", "Rap Price", "Rap Price($)", "Rap Rate", "RAP RTE", "Rap$", "RAP($)", "Rap-Price", "RAP.", "Rap.", "Price", "Rap.($)", "Rap/Price", "Rap_per_Crt", "RAP_PRICE", "Rapa", "Rapa Rate", "Rapa_Rate", "rapaport", "RAPAPORT_RATE", "RapaportPrice", "RapaRate", "RapDown", "Rape", "RapList", "RapNet Price", "rapnetcaratprice", "RapNetPrice", "RAPO", "RAPPLIST", "rapprice", "RapRat", "RapRate", "RapRice", "RapRte", "Rate", "repRate"],
      DISC: ["Disc", "%"," % Back"," % BELOW", "%Rap", "Asking Disc. %", "Back", "BACK %", "Back (-%)", "Back %", "Back -%", "Back%", "Base Off %", "Base Off%", "CBack", "DIC.", "DIS", "Dis %", "Dis%", "DIS.", "Disc", "Disc %", "Disc%", "Disc(%)", "DISC.", "Disc/Pre", "DISC_PER", "Disco%", "DISCOUNT", "Discount %","Discount % ??", "Discount%", "Discprct", "F disc", "Fair/Last Bid %", "Final %", "Final Disc%", "final_discount", "ListDisc%", "Net %", "New Rap%", "Off %", "Off%", "Offer Disc.(%)", "OffPer", "Price", "R.Dn", "Rap %", "RAP DIS", "Rap Disc", "Rap Disc %", "Rap Discount", "Rap%", "Rap.%", "RAP_DISCOUNT", "rap_per", "RapDis", "RapDown", "rapnet", "Rapnet", "Discount %", "RapNet Back", "Rapnet Discount", "Rapnet Discount%", "rapnetdiscount", "RapnetDiscountPercent", "RapOff", "RP Disc", "saleback", "SaleDis", "SaleDisc", "Selling Disc", "User Disc", "VDisc %"," WebsiteDiscount", "Rapdisc"],
    };

    const ResultSynonyms = {
      VALUE: ["value", "rapvalue", "rapaport value", "r.value", "val", "RapVlu"],
      NET_RATE: ["NetRate", "$ / Carat", "$/Carat", "$/CT", "$/CTS", "$/PC", "Asking Price", "askprice", "BACK P/Ct", "Base Rate", "Cash Price", "CashPrice", "CRate", "Ct/Price", "D.RAP PRICE", "DIS / CT", "Final Rate", "List$/Ct", "Net Rate", "NET_RATE", "P.CARAT", "P/CT", "P/CTS", "Per Crt $", "Per ct", "Per Ct $", "PerCarat", "PerCrt", "PerCts", "PPC", "PPC$", "Pr/Ct", "PRAP($)","PRI/CRT", "Price p.c", "Price $/cts", "Price / Crts", "Price Per Carat", "Price Per Crt", "Price Per Ct", "Price/Carat", "Price/Crt", "Price/Ct", "Price/Ct ($)", "Price/ct.", "Price/Cts", "Price/CTS $", "Price/Cts USD", "Price/Cts.", "PRICE_DOLLAR", "PRICE_PER_CARAT", "Price_Per_Crt", "PricePerCarat", "Rap @", "rap_prc", "RapNet Price", "RapNet Rate", "RATE", "Rate $/CT", "Rate / CT", "Rate ?", "Rate per carat as per Rapnet", "Rate($)", "RATE($/CT)", "Rate/Ct", "RP Price", "RTE", "SaleRate", "sales_price", "Selling Price", "User Price /Cts", "VALLUE", "WebsiteRate"],
      NET_VALUE: ["Net_Value", "$ Total", "amont", "AMOUNT", "Amount $", "Amount ?", "Amount US$", "Amount($)", "Amt", "Amt $", "Amt.", "askamount", "Asking Amount", "Back Total", "Base Amt", "CAmount", "DiscountPrice", "EST AMT", "F value", "F.Amt", "FINAL", "Final Amount", "Final Amt", "Final Amt IN $", "Final Price", "Final Value", "FINAL$", "final_amount", "FinalValue", "mspTotal", "Net", "NET VALLUE", "NET $", "Net Amt", "Net Amt($)", "Net Value", "NET_VALUE", "NetAmt", "Offer Value($)", "Rap US $", "Rapa Value", "RapNet Amount", "RapNet Price", "RP Tot$", "SaleAmt", "saledollorprice", "Stone Price", "Stone($)", "T AMT", "T Price", "T VALUE", "T. AMOUNT", "T.Amt", "Tot. Value", "Total", "TOTAL $", "Total $ as per Rapnet", "Total ($)", "TOTAL AMOUNT", "Total Amt", "Total Amt.", "Total Price", "Total$", "total_price", "TotalAmount", "TotalPrice", "TotalValue $", "User Total $", "VALUE_DOLLAR", "WebsiteAmount"],
    };

    // Find column indices (INPUT columns are required)
    function findColumnIndex(synonymsArray) {
      return headers.findIndex(header => 
        synonymsArray.some(synonym => 
          header === synonym.toLowerCase()
        )
      );
    }

    const wgtIndex = findColumnIndex(InputSynonyms.WGT);
    const rateIndex = findColumnIndex(InputSynonyms.RATE);
    const discIndex = findColumnIndex(InputSynonyms.DISC);

    if (wgtIndex === -1 || rateIndex === -1 || discIndex === -1) {
      console.error("Missing required columns (Weight, Rate, or Discount).");
      return;
    }

    // Find RESULT columns (optional)
    const valueIndex = findColumnIndex(ResultSynonyms.VALUE);
    const netRateIndex = findColumnIndex(ResultSynonyms.NET_RATE);
    const netValueIndex = findColumnIndex(ResultSynonyms.NET_VALUE);

    // Get column letters
    const wgtCol = getColumnLetter(wgtIndex);
    const rateCol = getColumnLetter(rateIndex);
    const discCol = getColumnLetter(discIndex);
    const valueCol = valueIndex !== -1 ? getColumnLetter(valueIndex) : null;
    const netRateCol = netRateIndex !== -1 ? getColumnLetter(netRateIndex) : null;
    const netValueCol = netValueIndex !== -1 ? getColumnLetter(netValueIndex) : null;

    // Apply formulas
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const rowNum = i + 1;

      if (valueCol) {
        data[i][valueIndex] = `=${wgtCol}${rowNum}*${rateCol}${rowNum}`;
      }

      if (netRateCol) {
        data[i][netRateIndex] = `=${rateCol}${rowNum}+((${rateCol}${rowNum}*${discCol}${rowNum})/100)`;
      }

      if (netRateCol && netValueCol) {
        data[i][netValueIndex] = `=${wgtCol}${rowNum}*${netRateCol}${rowNum}`;
      }
    }

    range.formulas = data;
    await context.sync();
    console.log("Formulas applied successfully!");
  }).catch(error => console.error("Error:", error));
}

// Helper function to convert column index to letter (A, B, ... AA, AB, etc.)
function getColumnLetter(index) {
  if (index < 0) return "";
  let columnLetter = "";
  let tempIndex = index;

  while (tempIndex >= 0) {
    columnLetter = String.fromCharCode((tempIndex % 26) + 65) + columnLetter;
    tempIndex = Math.floor(tempIndex / 26) - 1;
  }

  return columnLetter;
}

window.PutFormula = PutFormula;

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    currentWorksheet.load("name");
    await context.sync();

    const timestamp = new Date().getTime();
    const tableName = `ExpensesTable_${currentWorksheet.name.replace(/\s+/g, '_')}_${timestamp}`;

    const expensesTable = currentWorksheet.tables.add("A1:N1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.name = tableName;

    expensesTable.getHeaderRowRange().values = [["Shp#", "Color", "Clarity ??", "Cut", "Polish", "Symm", "FLName", "Lab", "Weight", "NEW RAP", "Disc",  "NetRate", "Amount $", "Value"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["Round", "E", "VVS1", "Good", "Good", "Good", "None", "G.I.A", "0.25", "5000", "-25", "", "", ""],
      ["Ht", "E", "VVS2-", "Ideal", "Ideal", "Ideal", "Non", "GIA", "0.98", "6000", "-27", "", "", ""],
      ["EM", "D", "SI1", "Ex", "Ex", "Ex", "MEDIUM", "HRD", "0.6", "5800", "-31", "", "", ""],
      ["Round", "XYZ", "SI2", "Gd", "Gd", "Gd", "None", "NCERT", "0.4", "5500", "-50", "", "", ""],
      ["EM", "D", "IF", "Excellent", "Ex", "Ex", "Non", "IGI", "1.25", "15000", "30", "", "", ""],
      ["TRI", "F YELLO", "LOUPE-CLEAN", "P", "POOR", "POOR", "SL", "NONE", "1.80", "8500", "-32", "", "", ""],
      ["HE", "MIX", "SI1", "FAIR", "F", "F", "ST-YL", "HRD", "0.6", "5800", "-31", "", "", ""],
      ["Princess",	"G",	"SI1",	"Excellent",	"Excellent",	"Excellent",	"STRONG",	"GIA",	"1.8",	"7800",	"8", "", "", ""],
      [ "Oval",	"H",	"SI2",	"Very Good", "	Very Good",	"Very Good",	"FNT",	"IGI",	"0.95",	"6800", "32", "", "", ""]	,
      [ "Cushion",	"I",	"I1",	"Good",	"Good",	"Good",	"SLIGHT",	"HRD",	"2.1",	"4500",	"-12", "", "", ""],
      [ "Emerald",	"D",	"VVS1",	"Excellent",	"Excellent",	"Excellent",	"None",	"GIA",	"1.1",	"13500",	"-25", "", "", ""],
      [ "OMB",	"E",	"VS1",	"Very Good",	"Very Good",	"Very Good",	"None",	"IGI",	"1.55",	"10500",	"-17", "", "", ""],
      ["Radiant",	"G",	"SI1",	"Excellent",	"Excellent",	"Excellent",	"FNT",	"GIA",	"1.7",	"6800",	"8", "", "", ""],
      ["Marquise",	"G",	"SI1",	"Poor",	"P",	"PR",	"MED",	"GIA",	"1.1",	"7500",	"-42", "", "", ""],
      ["Round",	"H",	"SI2",	"Very Good",	"Very Good",	"Very Good",	"NON",	"IGI",	"1.3",	"6500",	"-28", "", "", ""],
    ]);

    // Formatting
    expensesTable.columns.getItemAt(5).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
    console.log(`Table "${tableName}" created successfully on sheet "${currentWorksheet.name}"`);
  });
}


// Office.onReady(() => {
//     document.getElementById("openSortModal").addEventListener("click", openSortDialog);
// });

// function openSortDialog() {
//     Office.context.ui.displayDialogAsync(
//         "https://localhost:3000/dialog.html", // Change this URL based on your hosted add-in
//         { height: 50, width: 40, displayInIframe: true },
//         function (asyncResult) {
//             let dialog = asyncResult.value;

//             // Handle dialog messages
//             dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
//                 if (arg.message === "close") {
//                     dialog.close();
//                 }
//             });
//         }
//     );
// }

async function showCenterDialog() {
  try {
    await Excel.run(async (context) => {
      // Create dialog
      Office.context.ui.displayDialogAsync(
        'https://localhost:3001/dialog.html', // Replace with your dialog HTML URL
        {
          height: 50,  // Percentage of screen height
          width: 40,   // Percentage of screen width
          promptBeforeOpen: false,
          displayInIframe: true
        },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error(result.error.message);
          } else {
            // Store the dialog object
            const dialog = result.value;
            
            // Add event handlers
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
              console.log("Message received: " + args.message);
              dialog.close();
            });
            
            dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
              console.log("Dialog closed: " + args.error);
            });
          }
        }
      );
    });
  } catch (error) {
    console.error("Error showing dialog:", error);
  }
}

// Add this to your button click handler
document.getElementById('openCenterDialogBtn').addEventListener('click', () => {
  showCenterDialog();
});
