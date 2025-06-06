// ======================
// UI & Modal Management
// ======================
document.getElementById("openModal").addEventListener("click", function () {
  document.getElementById("customSortModal").style.display = "block";
  document.getElementById("modalOverlay").style.display = "block";
  loadPersistedColumns();
});

document.getElementById("closeModal").addEventListener("click", closeModal);
document.getElementById("modalOverlay").addEventListener("click", closeModal);

function closeModal() {
  document.getElementById("customSortModal").style.display = "none";
  document.getElementById("modalOverlay").style.display = "none";
}

const SortConfig = {
  // Maps standard values to their sort orders and synonyms
  orders: {
    Shape: {
      ROUND: { order: 1, synonyms: ["BRILLIANT", "RBC", "BR", "Round Brilliant", "ROUND", "RB", "Round"] },
      CUSHION: { order: 3, synonyms: ["CUSHION MODIFIED", "CM", "S.CUSHION", "SQ. CUSHION MODIFIED", "LONG CUSHION", "CUSHION MODIFIED BIALLIANT", "CM-4", "CM-4", "CM-4", "CMB", "CU", "CUSHION BRILLIANT", "CUSION", "CUSHION MBR", "CUS", "CUN", "CUSHION MODIFIED BRILLIANT", "CUSHION"] },
      HEART: { order: 6, synonyms: ["HE", "HB", "Ht", "S.HEART", "HT", "ROSE HEART", "HRT JODI", "Heart-P8-P8", "Heart"] },
      EMERALD: { order: 5, synonyms: ["EM", "Em", "E", "EMERALD 4STEP", "EMERALD", "ASYMEMERALD4S", "EM-HD", "EMRALD", "SQUARE EMERALD", "EMR", "Emerald Cut"] },
      PEAR: { order: 9, synonyms: ["PE", "SGP-100", "PB", "Pear Brilliant", "PEAR MODIFIED BRILLIANT"] },
      MQ: { order: 7, synonyms: ["M", "MQ", "MB", "S.MARQUISE", "MARQUISE", "MQ"] },
      OVAL: {  order: 8, synonyms: ["OB", "OVAL", "OV", "S.OVAL", "OMB", "Oval Brilliant", "OVL", "OVAL MODIFIED BRILLIANT"] },
      PRINCESS: { order: 10, synonyms: ["SMB", "BEZEL PRINCESS", "PRIN", "PRN", "PR", "PRINCES", "PRI", "Princess Cut", "PRINCESS", "PC"] },
      RADIAN: { order:  11, synonyms: ["LR", "RADIANTB", "RADIANT", "RN", "Rediant", "CSMB", "R", "CRMB", "CCRMB", "RA", "L-RD", "SQUARE RADIANT MODIFIED", "RADIANT MODIFIED BRILLIANT", "RADIANT MODIFIED", "LR_BRILLIANT", "LONG RADIANT", "SQ.RADIANT", "RADIANT", "SQ RADIANT"] },
      "SQ.EMERALD": { order: 12, synonyms: ["SEM", "SQ.EMERALD", "S.EMERALD", "SQ.EM", "SE", "SQ-EM", "SQ.EMERALD"] },
      BAGUETTE: { order: 14, synonyms: ["BT", "RC", "RSC", "BUG", "BU", "BUGGETTE", "BAGUETTE"] },
      TRIANGLE: { order: 16, synonyms: ["TR", "TS", "Triangular", "TRI"] },
      BRIOLETTE: { order: 17, synonyms: ["BCM", "BLT", "BRIOLETTE"] },
      "CUS.BRILLIANT": { order: 75, synonyms: ["CB", "CR", "CUS.BRILLIANT"] },
      MIX: { order: 94, synonyms: ["MIX"] },
      FANCY: { order: 95, synonyms: ["FANCY"] },
    },
    Color: {
      D: { order: 1, synonyms: ["D", "*D"] },
      E: { order: 3, synonyms: ["E"] },
      F: { order: 4, synonyms: ["F", "F YELLO"] },
      G: { order: 6, synonyms: ["G", "RARE WHITE G"] },
      H: { order: 7, synonyms: ["H", "H-", "WHITE (H)", "WHITE H"] },
      I: { order: 9, synonyms: ["I"] },
      J: { order: 10, synonyms: ["J"] },
      K: { order: 11, synonyms: ["K", "K, Faint Brown"] },
      L: { order: 12, synonyms: ["L", "L, Faint Brown"] },
      M: { order: 13, synonyms: ["M", "TINTED COLOUR M"] },
      N: { order: 14, synonyms: ["N"] },
      X: { order: 24, synonyms: ["X", "XYZ"] },
      Z: { order: 26, synonyms: ["Z"] },
      "P-Q": { order: 28, synonyms: ["P-Q"] },
      "S-T": { order: 30, synonyms: ["S-T"] },
      MIX: { order: 62, synonyms: ["MIX"] },
      FANCY: { order: 68, synonyms: ["FANCY"] },
    },
    Clarity: {
      IF: { order: 2, synonyms: ["IF", "LOUPE-CLEAN", "LC", "Internally Flawless"] },
      VVS1: { order: 3, synonyms: ["VVS1", "VVS1-", "vvs1", "VVS 1"] },
      VVS2: { order: 4, synonyms: ["VVS2", "VVS2-", "vvs2", "VVS 2"] },
      VS1: { order: 5, synonyms: ["VS1", "VS 1"] },
      VS2: { order: 6, synonyms: ["VS1", "VS 1"] },
      SI1: { order: 7, synonyms: ["SI1", "SI1-", "SI 1"] },
      SI2: { order: 8, synonyms: ["SI2", "SI2-", "SI 2"] },
      SI3: { order: 9, synonyms: ["SI3"] },
      I1: { order: 10, synonyms: ["I1"] },
      I2: { order: 11, synonyms: ["I1"] },
      MIX: { order: 16, synonyms: ["MIX"] },
    },
    Cut: {
      EX: { order: 2, synonyms: ["EXCELLENT", "Excellent", "EXC", "IDEAL", "Ideal", "Ex", "EX-2", "EX-1", "EX-3", "X1", "X2", "X3", "X4", "ex1", "ex2", "I", "Ex Ideal", "EX3"] },
      VG: { order: 3, synonyms: ["VERY GOOD", "V. GOOD", "VG", "VV", "VX", "V"] },
      GD: { order: 4, synonyms: ["GOOD", "Good", "GD", "G", "GV", "GX", "ex6", "G", "Gd"] },
      FR: { order: 5, synonyms: ["FAIR", "F", "FR", "FA", "F"] },
      PR: { order: 6, synonyms: ["P", "POOR", "PR", "PU"] },
    },
    Polish: {
      EX: { order: 2, synonyms: ["EXCELLENT", "Ex", "EXC", "IDEAL", "Ideal", "X1", "X2", "EX-2", "Ex Ideal"] },
      VG: { order: 3, synonyms: ["VERY GOOD", "VG", "V. GOOD"] },
      GD: { order: 4, synonyms: ["GOOD", "Good", "GD", "G", "Gd"] },
      FR: { order: 5, synonyms: ["FAIR", "FR", "F", "FA"] },
      PR: { order: 6, synonyms: ["P", "POOR", "PR"] },
    },
    Symm: {
      EX: { order: 2, synonyms: ["EXCELLENT", "Ex", "EXC", "IDEAL", "Ideal", "X1", "X2", "EX-2", "Ex Ideal"] },
      VG: { order: 3, synonyms: ["VERY GOOD", "VG", "V. GOOD"] },
      GD: { order: 4, synonyms: ["GOOD", "Good", "GD", "G", "Gd"] },
      FR: { order: 5, synonyms: ["FAIR", "FR", "F", "FA"] },
      PR: { order: 6, synonyms: ["P", "POOR", "PR"] },
    },
    Fls: {
      NONE: { order: 1, synonyms: ["N", "NO", "NIL", "FL0", "NN", "Non", "None"] },
      FNT: { order: 2, synonyms: ["FAINT", "FNT", "NEGLIGIBLE", "FL1", "FA"] },
      VSL: { order: 3, synonyms: ["VERY SLIGHT", "VSLG", "VSLT", "VSL", "VS"] },
      MED: { order: 4, synonyms: ["M", "MEDIUM", "MED", "FL2", "MD", "MB", "MEDIUM BLUE", "ME", "MD-BL", "MD-YL"] },
      SLT: { order: 5, synonyms: ["SL", "SLIGHT", "SLI", "SLT"] },
      STG: { order: 6, synonyms: ["STRONG", "STG", "ST", "FL3", "S", "STRONG BLUE", "STR", "STO", "ST-BL", "ST-WH", "ST-YL"] },
      VST: { order: 7, synonyms: ["VERY STRONG", "VST", "FL4", "VSTG", "VST-BL"] },
    },
    Lab: {
      GIA: { order: 1, synonyms: ["G.I.A", "GIA", "GA", "GIA"] },
      IGI: { order: 2, synonyms: ["IGI"] },
      HRD: { order: 3, synonyms: ["HRD"] },
      CGL: { order: 4, synonyms: ["CGL"] },
      AGS: { order: 5, synonyms: ["AGS"] },
      "NO-CERT": { order: 6, synonyms: ["NONE", "NON", "NC", "NONCERT", "NCERT", "NON CERT"] }
    },
  },

  // Defines which columns should be treated as numeric
  numericColumns: ["Weight", "Rate", "Disc", "Value", "NetRate", "Amount", "PRC", "DISCOUNT", "DISC"],

  // Maps synonyms to standard column names
  columnSynonyms: {
    Shape: ["Shape Name", "SHAPE#", "SHAP#", "Shp#", "Shape Desc", "Rap Shape", "rapShape", "Sh", "Sha", "SHAP", "SHAP.", "Shape", "Shape ??", "SHAPE.", "Shape_Code", "Shp", "Shp.", "Shp_Name"],
    Color: ["Colour", "Color (Long)", "Color C", "Col#", "CL.", "clours", "CLR", "Col", "Col.", "Color", "Color ??", "Color_Code", "ColorCode", "Colour", "Rap Color", "rapColor"],
    Clarity: ["Cl#", "Quality", "clar", "Cla#", "CAL", "Cal_Name", "CL", "Cl.", "Cla", "Cla.", "Clar", "Claratiy", "Clari", "Clarity", "Clarity ??", "Clarity_Code", "ClarityName", "CLERITY", "CLR", "Clr.", "Clrt", "CLRTY", "CTY", "Purity", "Qua", "Rap Clarity", "rapClarity"],
    Cut: ["Proportions", "CUTPROP", "CUT#", "Prop#", "Prop", "Ct", "CUT", "Cut Grade", "Cut Grade ??", "CUT.", "Cut_Code", "Cut_Grade", "CutGrade", "Final Cut", "Prop.", "PropCode"],
    Polish: ["POL#", "PL", "po", "POL", "POL.", "Polish", "Polish ??", "Polish_Code", "PolishName"],
    Symm: ["Symm", "SYM#", "SUM", "SYS", "Sy", "SYM", "Sym.", "SYMM", "Symmetry", "Symmetry ??", "Symmetry_Code", "Symmmetry", "SymName"],
    Fls: ["Flr", "FLRN", "FLUOR#", "Flour#", "FLRInt", "Fluore#", "Flor#", "Flo#", "FL", "FL.", "FLName", "FLO", "flor", "FLOR.", "Flore", "Floro", "FLORO.", "Florosence", "Flors", "Flou", "Flour", "Flour.", "Flourence", "FLOURESCENSE", "FLOURESENCE", "Flr", "flr_intensity", "Flrcnt", "FLRInt", "FlrIntens", "Fls", "FLS.", "FLU", "fluo", "Fluo Int", "Fluo.", "Fluor", "Fluor.", "Fluores..", "Fluorescence", "Fluorescence ??", "Fluorescence Intensity", "Fluorescence_Code", "Fluorescence_Intensity", "FluorescenceColor", "FluorescenceIntensity", "Fluorescense", "Fluorescent", "FLUORS", "Flur", "Fluro", "FLURO."],
    Lab: ["Cer", "Cert", "Cert By", "Cert From", "Cert Name", "CERT.", "CERTI", "CERTI_NAME", "CERTIFICATE", "CertName", "CR_Name", "Crt", "Lab", "Lab ??", "Lab Name", "Lab_Code", "report"],
    Weight: ["TOTAL CTS","TotalCts", "Weight R","weigh", "Cts#", "SIZE#","Wt#", "Car", "Cara", "Carat", "CARATS", "Crt", "Crts", "CRTWT", "CT", "Ct.", "Cts", "Cts.", "POLISE" ,"CT" ,"Size", "SIZE." ,"Weight", "Weight ??", "Wgt" ,"WHT.", "WT", "Wt."],
    Rate: ["BaseRate", "Disc Price"," Full Rap Price", "List", "List Price", "List Price ????", "List Rate", "LiveRAP", "NEW RAP", "Orap", "price", "R.PRICE", "Rap", "Rap $", "Rap $/CT", "Rap List", "Rap Price", "Rap Price($)", "Rap Rate", "RAP RTE", "Rap$", "RAP($)", "Rap-Price", "RAP.", "Rap.", "Price", "Rap.($)", "Rap/Price", "Rap_per_Crt", "RAP_PRICE", "Rapa", "Rapa Rate", "Rapa_Rate", "rapaport", "RAPAPORT_RATE", "RapaportPrice", "RapaRate", "RapDown", "Rape", "RapList", "RapNet Price", "rapnetcaratprice", "RapNetPrice", "RAPO", "RAPPLIST", "rapprice", "RapRat", "RapRate", "RapRice", "RapRte", "Rate", "repRate"],
    Disc: ["%"," % Back"," % BELOW", "%Rap", "Asking Disc. %", "Back", "BACK %", "Back (-%)", "Back %", "Back -%", "Back%", "Base Off %", "Base Off%", "CBack", "DIC.", "DIS", "Dis %", "Dis%", "DIS.", "Disc", "Disc %", "Disc%", "Disc(%)", "DISC.", "Disc/Pre", "DISC_PER", "Disco%", "DISCOUNT", "Discount %","Discount % ??", "Discount%", "Discprct", "F disc", "Fair/Last Bid %", "Final %", "Final Disc%", "final_discount", "ListDisc%", "Net %", "New Rap%", "Off %", "Off%", "Offer Disc.(%)", "OffPer", "Price", "R.Dn", "Rap %", "RAP DIS", "Rap Disc", "Rap Disc %", "Rap Discount", "Rap%", "Rap.%", "RAP_DISCOUNT", "rap_per", "RapDis", "RapDown", "rapnet", "Rapnet", "Discount %", "RapNet Back", "Rapnet Discount", "Rapnet Discount%", "rapnetdiscount", "RapnetDiscountPercent", "RapOff", "RP Disc", "saleback", "SaleDis", "SaleDisc", "Selling Disc", "User Disc", "VDisc %"," WebsiteDiscount", "Rapdisc"],
    Value: ["value", "rapvalue", "rapaport value", "r.value", "val", "RapVlu"],
    NetRate: ["net_rate", "$ / Carat", "$/Carat", "$/CT", "$/CTS", "$/PC", "Asking Price", "askprice", "BACK P/Ct", "Base Rate", "Cash Price", "CashPrice", "CRate", "Ct/Price", "D.RAP PRICE", "DIS / CT", "Final Rate", "List$/Ct", "Net Rate", "NET_RATE", "P.CARAT", "P/CT", "P/CTS", "Per Crt $", "Per ct", "Per Ct $", "PerCarat", "PerCrt", "PerCts", "PPC", "PPC$", "Pr/Ct", "PRAP($)","PRI/CRT", "Price p.c", "Price $/cts", "Price / Crts", "Price Per Carat", "Price Per Crt", "Price Per Ct", "Price/Carat", "Price/Crt", "Price/Ct", "Price/Ct ($)", "Price/ct.", "Price/Cts", "Price/CTS $", "Price/Cts USD", "Price/Cts.", "PRICE_DOLLAR", "PRICE_PER_CARAT", "Price_Per_Crt", "PricePerCarat", "Rap @", "rap_prc", "RapNet Price", "RapNet Rate", "RATE", "Rate $/CT", "Rate / CT", "Rate ?", "Rate per carat as per Rapnet", "Rate($)", "RATE($/CT)", "Rate/Ct", "RP Price", "RTE", "SaleRate", "sales_price", "Selling Price", "User Price /Cts", "VALLUE", "WebsiteRate"],
    Amount: ["net_value", "$ Total", "amont", "AMOUNT", "Amount $", "Amount ?", "Amount US$", "Amount($)", "Amt", "Amt $", "Amt.", "askamount", "Asking Amount", "Back Total", "Base Amt", "CAmount", "DiscountPrice", "EST AMT", "F value", "F.Amt", "FINAL", "Final Amount", "Final Amt", "Final Amt IN $", "Final Price", "Final Value", "FINAL$", "final_amount", "FinalValue", "mspTotal", "Net", "NET VALLUE", "NET $", "Net Amt", "Net Amt($)", "Net Value", "NET_VALUE", "NetAmt", "Offer Value($)", "Rap US $", "Rapa Value", "RapNet Amount", "RapNet Price", "RP Tot$", "SaleAmt", "saledollorprice", "Stone Price", "Stone($)", "T AMT", "T Price", "T VALUE", "T. AMOUNT", "T.Amt", "Tot. Value", "Total", "TOTAL $", "Total $ as per Rapnet", "Total ($)", "TOTAL AMOUNT", "Total Amt", "Total Amt.", "Total Price", "Total$", "total_price", "TotalAmount", "TotalPrice", "TotalValue $", "User Total $", "VALUE_DOLLAR", "WebsiteAmount"]
  },

  // Gets the standard name for a column
  getStandardColumnName: function(columnName) {
    if (!columnName) return null;
    const cleanName = columnName.trim().toUpperCase();
    for (const [stdName, synonyms] of Object.entries(this.columnSynonyms)) {
      if (stdName.toUpperCase() === cleanName) return stdName;
      if (synonyms.some(syn => syn.toUpperCase() === cleanName)) return stdName;
    }
    return columnName;
  },

  // Gets the standard value and sort order for a data value
  getStandardValueInfo: function (value, columnName) {
    const stdColumnName = this.getStandardColumnName(columnName);
    if (!stdColumnName || !this.orders[stdColumnName]) return null;

    const strValue = String(value).trim().toUpperCase();
    const category = this.orders[stdColumnName];

    // Check for direct match with standard values
    for (const [stdVal, info] of Object.entries(category)) {
      if (stdVal.toUpperCase() === strValue) {
        return { standardValue: stdVal, order: info.order };
      }
    }

    // Check synonyms
    for (const [stdVal, info] of Object.entries(category)) {
      if (info.synonyms.some(syn => syn.toUpperCase() === strValue)) {
        return { standardValue: stdVal, order: info.order };
      }
    }

    return null;
  },
};

// modal & Dropdown Content
let allColumns = []; // Stores all available columns from the sheet

document.getElementById("openModal").addEventListener("click", function () {
  document.getElementById("customSortModal").style.display = "block";
  document.getElementById("modalOverlay").style.display = "block";
  populateColumnDropdown();
});

document.getElementById("closeModal").addEventListener("click", closeModal);
document.getElementById("modalOverlay").addEventListener("click", closeModal);

function closeModal() {
  document.getElementById("customSortModal").style.display = "none";
  document.getElementById("modalOverlay").style.display = "none";
}

async function populateColumnDropdown() {
  const dropdown = document.getElementById("dropdown1");
  dropdown.innerHTML = '<option value="">Select Column</option>';

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();

      if (range.values && range.values.length > 0) {
        // Map worksheet columns to their standardized names
        const worksheetColumns = range.values[0];
        const standardizedColumns = worksheetColumns.map(col => {
          // Find the standardized name for this column
          for (const [standardName, synonyms] of Object.entries(SortConfig.columnSynonyms)) {
            if (col === standardName || synonyms.includes(col)) {
              return standardName;
            }
          }
          return col; // Return as-is if no synonym match
        });

        // Store both original and standardized columns
        allColumns = {
          original: worksheetColumns,
          standardized: [...new Set(standardizedColumns)] // Remove duplicates
        };

        refreshDropdownOptions();
      }
    });
  } catch (error) {
    console.error("Error loading columns:", error);
  }
}

function refreshDropdownOptions() {
  const dropdown = document.getElementById("dropdown1");
  const currentItems = [...document.getElementById("selectedColumns").rows]
    .map(row => row.cells[1].textContent);

  dropdown.innerHTML = '<option value="">Select Column</option>';

  allColumns.standardized.forEach((standardName, index) => {
    if (standardName) {
      const option = document.createElement("option");
      option.value = standardName; // Store standardized name as value
      option.textContent = standardName; // Show standardized name

      // Find matching original column name
      const originalName = allColumns.original.find((col, i) => {
        const synonyms = SortConfig.columnSynonyms[standardName] || [];
        return col === standardName || synonyms.includes(col);
      }) || standardName;

      option.dataset.originalName = originalName; // Store original name in data attribute

      // Disable if already added
      option.disabled = currentItems.includes(standardName) && 
                       !removedColumns.includes(standardName);

      if (currentItems.includes(standardName)) {
        option.style.color = "#ff0000";
        option.title = "Already added";
      } else if (removedColumns.includes(standardName)) {
        option.title = "Removed - can re-add";
      }

      dropdown.appendChild(option);
    }
  });
}

// ==============================================
// 3. UPDATED SAVE FUNCTION (Handles new order)
// ==============================================

function createTableRow(column, index) {
  return `
    <td>${index + 1}</td>
    <td>${column.columnName}</td>
    <td>
      <select>
        <option value="ASC"  ${column.sortOrder === "ASC" ? "selected" : ""}>ASC</option>
        <option value="DESC" ${column.sortOrder === "DESC" ? "selected" : ""}>DESC</option>
      </select>
    </td>
    <td class="action-buttons">
      <div class="move-btn-group">
        <button class="move-up-btn" onclick="moveColumnUp(this)">↑</button>
        <button class="move-down-btn" onclick="moveColumnDown(this)">↓</button>
      </div>
      <button class="close-row-btn" onclick="removeColumn(this)">✕</button>
    </td>
  `;
}

function moveColumnUp(button) {
  const row = button.closest("tr");
  const prevRow = row.previousElementSibling;
  if (prevRow) {
    row.parentNode.insertBefore(row, prevRow);
    updateRowNumbers();
  }
}

function moveColumnDown(button) {
  const row = button.closest("tr");
  const nextRow = row.nextElementSibling;
  if (nextRow) {
    row.parentNode.insertBefore(nextRow, row);
    updateRowNumbers();
  }
}

function updateRowNumbers() {
  const rows = document.querySelectorAll("#selectedColumns tr");
  rows.forEach((row, index) => {
    row.cells[0].textContent = index + 1;
  });
}

// Initialize when modal opens
document.addEventListener("DOMContentLoaded", () => {
  loadPersistedColumns();
});

let sortSettingsDialog = null;

function initDialog() {
  if (!sortSettingsDialog) {
    const dialogHtml = `
      <div class="modal-overlay2" id="sortSettingsOverlay" style="display:none">
        <div class="save-modal2">
          <div class="modal-header2">
            <h3>Sort Settings Saved</h3>
          </div>
          <div class="modal-body2">
            <p id="sortSettingsMessage"></p>
          </div>
          <div class="modal-footer2">
            <button class="modal-ok-btn2" onclick="hideDialog()">OK</button>
          </div>
        </div>
      </div>
    `;
    document.body.insertAdjacentHTML("beforeend", dialogHtml);
    sortSettingsDialog = {
      overlay: document.getElementById("sortSettingsOverlay"),
      message: document.getElementById("sortSettingsMessage"),
    };
  }
}

// DIALOG FUNCTIONS (Reuses same DOM element)

function showDialog(message) {
  if (!sortSettingsDialog) initDialog();
  sortSettingsDialog.message.textContent = message;
  sortSettingsDialog.overlay.style.display = "flex";
  setTimeout(() => sortSettingsDialog.overlay.classList.add("active"), 10);
}

function hideDialog() {
  if (sortSettingsDialog) {
    sortSettingsDialog.overlay.classList.remove("active");
    setTimeout(() => (sortSettingsDialog.overlay.style.display = "none"), 300);
  }
}

// XML STORAGE FUNCTIONS

const SORT_SETTINGS_FILE = "sort_settings.xml";

function escapeXml(unsafe) {
  if (!unsafe) return "";
  return unsafe
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// Generate XML from columns array
function generateSortSettingsXml(columns) {
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<SortSettings>\n';
  columns.forEach((col) => {
    xml += `  <Column>\n`;
    xml += `    <Name>${escapeXml(col.columnName)}</Name>\n`;
    xml += `    <Order>${escapeXml(col.sortOrder)}</Order>\n`;
    xml += `  </Column>\n`;
  });
  xml += "</SortSettings>";
  return xml;
}

// Parse XML to columns array
function parseXmlToColumns(xmlText) {
  try {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, "text/xml");
    const columns = [];

    xmlDoc.querySelectorAll("Column").forEach((colNode) => {
      columns.push({
        columnName: colNode.querySelector("Name").textContent,
        sortOrder: colNode.querySelector("Order").textContent,
      });
    });

    return columns;
  } catch (error) {
    console.error("XML parse error:", error);
    return [];
  }
}

async function saveToXmlFile(xmlContent) {
  try {
    const db = await initializeDB();
    const transaction = db.transaction("SortSettings", "readwrite");
    const store = transaction.objectStore("SortSettings");
    store.put(xmlContent, SORT_SETTINGS_FILE);
  } catch (error) {
    console.error("Save error:", error);
    throw error;
  }
}

async function loadFromXmlFile() {
  try {
    const db = await initializeDB();
    return new Promise((resolve) => {
      const transaction = db.transaction("SortSettings", "readonly");
      const store = transaction.objectStore("SortSettings");
      const request = store.get(SORT_SETTINGS_FILE);

      request.onsuccess = () => {
        resolve(request.result || '<?xml version="1.0"?><SortSettings></SortSettings>');
      };

      request.onerror = () => {
        resolve('<?xml version="1.0"?><SortSettings></SortSettings>');
      };
    });
  } catch (error) {
    console.error("Load error:", error);
    return '<?xml version="1.0"?><SortSettings></SortSettings>';
  }
}

// UPDATED SAVE COLUMN FUNCTION

// async function saveSortSettingsPermanently() {
//   try {
//     const table = document.getElementById("selectedColumns");
//     const columns = Array.from(table.rows).map(row => ({
//       columnName: row.cells[1].textContent.trim(),
//       sortOrder: row.cells[2].querySelector("select").value
//     }));

//     localStorage.setItem("customSortColumns", JSON.stringify(columns));
//     removedColumns = []; // Clear removed columns on save
//     refreshDropdownOptions(); // Update dropdown state

//     showDialog(`${columns.length} column${columns.length !== 1 ? 's' : ''} saved successfully.`);

//   } catch (error) {
//     showDialog("Failed to save sort settings");
//     console.error("Save error:", error);
//   }
// }

// Initialize dialog when add-in loads
Office.onReady(() => initDialog());

/**
 * Updates the sequence numbers in the first column
 */
function updateSortOrderNumbers() {
  const table = document.getElementById("selectedColumns");
  Array.from(table.rows).forEach((row, index) => {
    row.cells[0].textContent = index + 1;
  });
}

/**
 * Saves the current sort settings
 */
// function saveCurrentSortSettings() {
//   const settings = [];
//   const table = document.getElementById("selectedColumns");

//   Array.from(table.rows).forEach(row => {
//     settings.push({
//       columnName: row.cells[1].textContent,
//       sortOrder: row.cells[2].querySelector("select").value
//     });
//   });

//   Office.context.document.settings.set("sortSettings", JSON.stringify(settings));
//   Office.context.document.settings.saveAsync();
// }

// 2. LOAD FUNCTION (With Close Buttons)

// function loadPersistedColumns() {
//   const savedColumns = JSON.parse(localStorage.getItem("customSortColumns") || "[]");
//   const tbody = document.getElementById("selectedColumns");
//   tbody.innerHTML = "";

//   savedColumns.forEach((col, index) => {
//     const row = tbody.insertRow();
//     row.innerHTML = `
//       <td>${index + 1}</td>
//       <td>${col.columnName}</td>
//       <td>
//         <select onchange="saveSortSettingsPermanently()">
//           <option value="ASC" ${col.sortOrder === "ASC" ? "selected" : ""}>ASC</option>
//           <option value="DESC" ${col.sortOrder === "DESC" ? "selected" : ""}>DESC</option>
//         </select>
//       </td>
//       <td><button class="close-row-btn" onclick="removeColumn(this)">✕</button></td>
//     `;
//   });
// }

// UPGRADED MAIN FUNCTIONS XML STORAGE

async function saveSortSettingsPermanently() {
  try {
    const table = document.getElementById("selectedColumns");
    const columns = Array.from(table.rows).map((row) => ({
      columnName: row.cells[1].textContent.trim(),
      sortOrder: row.cells[2].querySelector("select").value,
    }));

    const xml = generateSortSettingsXml(columns);
    await saveToXmlFile(xml);

    removedColumns = [];
    refreshDropdownOptions();

    showDialog(`${columns.length} column${columns.length !== 1 ? "s" : ""} saved successfully.`);
  } catch (error) {
    showDialog("Failed to save sort settings");
    console.error("Save error:", error);
  }
}

async function loadPersistedColumns() {
  try {
    const xml = await loadFromXmlFile();
    const savedColumns = parseXmlToColumns(xml);
    const tbody = document.getElementById("selectedColumns");

    if (!tbody) {
      console.error("Table body not found");
      return;
    }

    tbody.innerHTML = "";
    savedColumns.forEach((col, index) => {
      const row = tbody.insertRow();
      row.innerHTML = createTableRow(col, index);
    });
  } catch (error) {
    console.error("Load error:", error);
    showDialog("Error loading saved columns");
  }
}

let removedColumns = [];

// ========== FIXED CORE FUNCTIONS ==========
async function removeColumn(button) {
  try {
    const row = button.closest("tr");
    if (!row) return;

    const columnName = row.cells[1]?.textContent?.trim();
    if (!columnName) return;

    removedColumns.push(columnName);

    // Update XML storage
    const xml = await loadFromXmlFile();
    const columns = parseXmlToColumns(xml).filter((col) => col.columnName !== columnName);
    await saveToXmlFile(generateSortSettingsXml(columns));

    row.remove();
    updateRowNumbers();
    if (typeof refreshDropdownOptions === "function") {
      refreshDropdownOptions();
    }
  } catch (error) {
    console.error("Column removal error:", error);
    showDialog("Error removing column");
  }
}

let sortColumns = [];

// ADD NEW COLUMN FUNCTION

function addColumn() {
  const dropdown = document.getElementById("dropdown1");
  const selectedColumn = dropdown.value;
  if (!selectedColumn) return;

  const tbody = document.getElementById("selectedColumns");

  removedColumns = removedColumns.filter((col) => col !== selectedColumn);

  // Check if column already exists
  const existingColumns = Array.from(tbody.rows).map((row) => row.cells[1].textContent.trim());
  if (existingColumns.includes(selectedColumn)) {
    console.log("Column already added!", "warning");
    return;
  }

  // Add new row
  const newRow = tbody.insertRow();
  newRow.innerHTML = `
    <td>${tbody.rows.length}</td>
    <td>${selectedColumn}</td>
    <td>
      <select>
        <option value="ASC">ASC</option>
        <option value="DESC">DESC</option>
      </select>
    </td>
    <td class="action-buttons"> 
      <div class="move-btn-group">
        <button class="move-up-btn" onclick="moveColumnUp(this)">↑</button>
        <button class="move-down-btn" onclick="moveColumnDown(this)">↓</button>
      </div>
      <button class="close-row-btn" onclick="removeColumn(this)">✕</button>
    </td>
  `;

  updateRowNumbers(); // Update numbering without saving
}

// Main Sorting Logic Function
/**
 * Applies the sorting to the Excel sheet
 */
async function applySorting() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load(["values", "address"]);
      await context.sync();

      if (!range.values || range.values.length < 2) {
        alert("No data to sort!");
        return;
      }

      const savedColumns = JSON.parse(localStorage.getItem("customSortColumns") || "[]");

      if (savedColumns.length === 0) {
        alert("No saved sort columns found!");
        return;
      }

      // Apply all saved columns
      sortColumns.forEach((col) => {
        range.sort.apply(
          [
            {
              key: col.column,
              ascending: col.order === "ASC",
            },
          ],
          true
        );
      });

      loadPersistedColumns();

      const headers = range.values[0];
      const data = range.values.slice(1);

      const currentSortSettings = getCurrentSortSettings();

      const sortedData = multiColumnSort(data, headers, currentSortSettings);

      sheet.getRangeByIndexes(1, 0, sortedData.length, sortedData[0].length).values = sortedData;
      await context.sync();

      alert("Data re-sorted successfully with current settings!");
      closeModal();
    });
  } catch (error) {
    console.error("Sorting failed:", error);
    alert("Sorting failed. See console for details.");
  }
}

/**
 * Gets the CURRENT sort settings from the UI
 */
function getCurrentSortSettings() {
  const settings = [];
  const table = document.getElementById("selectedColumns");

  Array.from(table.rows).forEach((row) => {
    settings.push({
      columnName: row.cells[1].textContent,
      sortOrder: row.cells[2].querySelector("select").value,
    });
  });

  return settings;
}

/**
 * Gets the original column name from the worksheet that matches the standard name
 */
function getOriginalColumnName(standardName) {
  if (!allColumns?.original) return standardName;
  
  // Check exact match first
  if (allColumns.original.includes(standardName)) {
    return standardName;
  }
  
  // Check synonyms
  const synonyms = SortConfig.columnSynonyms[standardName] || [];
  for (const synonym of synonyms) {
    if (allColumns.original.includes(synonym)) {
      return synonym;
    }
  }
  
  return standardName; // Fallback
}

/**
 * Updated multiColumnSort to handle synonyms
 */
function multiColumnSort(data, headers, customSortColumns) {
  return [...data].sort((rowA, rowB) => {
    for (const setting of customSortColumns) {
      const { columnName, sortOrder } = setting;
      
      // Find the actual column name in the worksheet
      const originalColName = getOriginalColumnName(columnName);
      const colIndex = headers.findIndex(h => h === originalColName);

      if (colIndex === -1) continue;

      const valueA = rowA[colIndex];
      const valueB = rowB[colIndex];

      if (!valueA && !valueB) continue;

      const compareResult = compareValues(valueA, valueB, columnName);

      if (compareResult !== 0) {
        return sortOrder === "ASC" ? compareResult : -compareResult;
      }
    }
    return 0;
  });
}

/**
 * Compares two values based on column type
 */
function compareValues(a, b, columnName) {
  if (!a) return 1;
  if (!b) return -1;

  const stdColumnName = SortConfig.getStandardColumnName(columnName);

  if (
    SortConfig.numericColumns.includes(stdColumnName) ||
    SortConfig.numericColumns.some((col) => SortConfig.columnSynonyms[col]?.includes(stdColumnName))
  ) {
    const numA = parseFloat(a) || 0;
    const numB = parseFloat(b) || 0;
    return numA - numB;
  }

  // Get standardized values and their sort orders
  const valAInfo = SortConfig.getStandardValueInfo(a, stdColumnName);
  const valBInfo = SortConfig.getStandardValueInfo(b, stdColumnName);

  const orderA = valAInfo ? valAInfo.order : 9999;
  const orderB = valBInfo ? valBInfo.order : 9999;

  return orderA - orderB;
}

async function initializeDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("ExcelSortSettingsDB", 1);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains("SortSettings")) {
        db.createObjectStore("SortSettings");
      }
    };

    request.onsuccess = (event) => {
      resolve(event.target.result);
    };

    request.onerror = (event) => {
      console.error("IndexedDB error:", event.target.error);
      reject(event.target.error);
    };
  });
}
