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

function handleNetCalculations() {
  console.log("Calculating Values, NetRate, and NetValue...");

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("rowCount, columnCount, values");
    await context.sync();

    const rowCount = range.rowCount;

    for (let i = 1; i < rowCount; i++) {
      const valueCell = sheet.getCell(i, 3);
      const netRateCell = sheet.getCell(i, 4);
      const netValueCell = sheet.getCell(i, 5);

      // Formulas
      valueCell.formulas = [[`=A${i + 1}*B${i + 1}`]];

      netRateCell.formulas = [[`=B${i + 1}+(B${i + 1}*C${i + 1}/100)`]];

      netValueCell.formulas = [[`=A${i + 1}*E${i + 1}`]];
    }

    const spacingRow = rowCount;
    const lastRow = rowCount + 1;
    sheet.getRange(`A${spacingRow + 1}:F${spacingRow + 1}`).values = [["", "", "", "", "", ""]];

    sheet.getCell(lastRow, 0).formulas = [[`=SUM(A2:A${rowCount})`]];
    sheet.getCell(lastRow, 3).formulas = [[`=SUM(D2:D${rowCount})`]];
    sheet.getCell(lastRow, 5).formulas = [[`=SUM(F2:F${rowCount})`]];

    // Division formulas
    sheet.getCell(lastRow, 1).formulas = [[`=D${lastRow + 1}/A${lastRow + 1}`]];
    sheet.getCell(lastRow, 4).formulas = [[`=F${lastRow + 1}/A${lastRow + 1}`]];

    const lastRowRange = sheet.getRange(`A${lastRow + 1}:F${lastRow + 1}`);
    lastRowRange.format.fill.color = "yellow";

    await context.sync();
    console.log("All calculations, spacing, and styling applied successfully.");
  }).catch((error) => {
    console.error("Error in handleNetCalculations: ", error);
  });
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
      NET_RATE: ["net_rate", "$ / Carat", "$/Carat", "$/CT", "$/CTS", "$/PC", "Asking Price", "askprice", "BACK P/Ct", "Base Rate", "Cash Price", "CashPrice", "CRate", "Ct/Price", "D.RAP PRICE", "DIS / CT", "Final Rate", "List$/Ct", "Net Rate", "NET_RATE", "P.CARAT", "P/CT", "P/CTS", "Per Crt $", "Per ct", "Per Ct $", "PerCarat", "PerCrt", "PerCts", "PPC", "PPC$", "Pr/Ct", "PRAP($)","PRI/CRT", "Price p.c", "Price $/cts", "Price / Crts", "Price Per Carat", "Price Per Crt", "Price Per Ct", "Price/Carat", "Price/Crt", "Price/Ct", "Price/Ct ($)", "Price/ct.", "Price/Cts", "Price/CTS $", "Price/Cts USD", "Price/Cts.", "PRICE_DOLLAR", "PRICE_PER_CARAT", "Price_Per_Crt", "PricePerCarat", "Rap @", "rap_prc", "RapNet Price", "RapNet Rate", "RATE", "Rate $/CT", "Rate / CT", "Rate ?", "Rate per carat as per Rapnet", "Rate($)", "RATE($/CT)", "Rate/Ct", "RP Price", "RTE", "SaleRate", "sales_price", "Selling Price", "User Price /Cts", "VALLUE", "WebsiteRate"],
      NET_VALUE: ["net_value", "$ Total", "amont", "AMOUNT", "Amount $", "Amount ?", "Amount US$", "Amount($)", "Amt", "Amt $", "Amt.", "askamount", "Asking Amount", "Back Total", "Base Amt", "CAmount", "DiscountPrice", "EST AMT", "F value", "F.Amt", "FINAL", "Final Amount", "Final Amt", "Final Amt IN $", "Final Price", "Final Value", "FINAL$", "final_amount", "FinalValue", "mspTotal", "Net", "NET VALLUE", "NET $", "Net Amt", "Net Amt($)", "Net Value", "NET_VALUE", "NetAmt", "Offer Value($)", "Rap US $", "Rapa Value", "RapNet Amount", "RapNet Price", "RP Tot$", "SaleAmt", "saledollorprice", "Stone Price", "Stone($)", "T AMT", "T Price", "T VALUE", "T. AMOUNT", "T.Amt", "Tot. Value", "Total", "TOTAL $", "Total $ as per Rapnet", "Total ($)", "TOTAL AMOUNT", "Total Amt", "Total Amt.", "Total Price", "Total$", "total_price", "TotalAmount", "TotalPrice", "TotalValue $", "User Total $", "VALUE_DOLLAR", "WebsiteAmount"],
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

    expensesTable.getHeaderRowRange().values = [["Shp#", "Color", "Clarity ??", "Cut", "Polish", "Symm", "FLName", "Lab", "Weight", "NEW RAP", "Disc", "value", "net_rate", "Amount $"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["Round", "E", "VVS1", "Good", "Good", "Good", "None", "G.I.A", "0.25", "5000", "-25", "", "", ""],
      ["Ht", "E", "VVS2-", "Ideal", "Ideal", "Ideal", "Non", "GIA", "0.98", "6000", "-27", "", "", ""],
      ["EM", "D", "SI1", "Ex", "Ex", "Ex", "MEDIUM", "HRD", "0.6", "5800", "-31", "", "", ""],
      ["Round", "XYZ", "SI2", "Gd", "Gd", "Gd", "None", "NCERT", "0.4", "5500", "-50", "", "", ""],
      ["EM", "D", "IF", "Excellent", "Ex", "Ex", "Non", "IGI", "1.25", "15000", "30", "", "", ""],
      ["TRI", "F YELLO", "LOUPE-CLEAN", "P", "POOR", "POOR", "SL", "NONE", "1.80", "8500", "-32", "", "", ""],
      ["HE", "MIX", "SI1", "FAIR", "F", "F", "ST-YL", "HRD", "0.6", "5800", "-31", "", "", ""],
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
