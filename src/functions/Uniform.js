/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */

function UniformData() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    if (!range.values || range.values.length === 0) {
      console.error("No data found in the selected range.");
      return;
    }

    const originalHeaders = range.values[0];

    // 1. Create a map of which columns to standardize based on ColumnSynonyms
    const columnsToStandardize = originalHeaders.map(header => {
      if (typeof header !== 'string') return null;

      const cleanHeader = header.trim().toUpperCase();

      // Find which category this header belongs to (if any)
      for (const [category, synonyms] of Object.entries(ColumnSynonyms)) {
        if (category.toUpperCase() === cleanHeader || synonyms.some(syn => syn.toUpperCase() === cleanHeader)) {
          return category; // This column should be standardized
        }
      }
      return null; // No standardization needed for this column
    });

    // 2. Standardize only the data cells in columns that matched ColumnSynonyms
    const standardizedData = range.values.map((row, rowIndex) => {
      if (rowIndex === 0) return row; // Keep original headers

      return row.map((cell, colIndex) => {
        const category = columnsToStandardize[colIndex];
        if (!category || typeof cell !== 'string') return cell;

        const cleanCell = cell.trim().toUpperCase();

        // Find matching uniform word in Synonyms
        for (const [uniformWord, synonyms] of Object.entries(Synonyms[category])) {
          if (uniformWord.toUpperCase() === cleanCell || synonyms.some((syn) => syn.toUpperCase() === cleanCell)) {
            return uniformWord;
          }
        }
        return cell; // Return original if no match found
      });
    });

    // 3. Write back the data (with original headers)
    range.values = standardizedData;
    await context.sync();
    console.log("Uniform Data Working");
  }).catch(error => {
    console.error("Error in UniformData:", error);
    throw error;
  });
}

// Combined synonyms dictionaries
const Synonyms = {
  Shape: {
    ROUND: ["BRILLIANT", "RBC", "BR", "Round Brilliant", "ROUND", "RB", "Round"],
    CUSHION: ["CUSHION MODIFIED", "CM", "S.CUSHION", "SQ. CUSHION MODIFIED", "LONG CUSHION", "CUSHION MODIFIED BIALLIANT", "CM-4", "CM-4", "CM-4", "CMB", "CU", "CUSHION BRILLIANT", "CUSION", "CUSHION MBR", "CUS", "CUN", "CUSHION MODIFIED BRILLIANT", "CUSHION"],
    HEART: ["HE", "HB", "Ht", "S.HEART", "HT", "ROSE HEART", "HRT JODI", "Heart-P8-P8", "Heart"],
    EMERALD: ["EM", "Em", "E", "EMERALD 4STEP", "EMERALD", "ASYMEMERALD4S", "EM-HD", "EMRALD", "SQUARE EMERALD", "EMR", "Emerald Cut"],
    PEAR: ["PE", "SGP-100", "PB", "Pear Brilliant", "PEAR MODIFIED BRILLIANT"],
    OVAL: ["OB", "OVAL", "OV", "S.OVAL", "OMB", "Oval Brilliant", "OVL", "OVAL MODIFIED BRILLIANT"],
    TRIANGLE: ["TR", "TS", "Triangular", "TRI"],
    MQ: ["M", "MQ", "MB", "S.MARQUISE", "MARQUISE", "MQ"],
    PRINCESS: ["SMB", "BEZEL PRINCESS", "PRIN", "PRN", "PR", "PRINCES", "PRI", "Princess Cut", "PRINCESS", "PC"] ,
    RADIAN: ["LR", "RADIANTB", "RADIANT", "RN", "Rediant", "CSMB", "R", "CRMB", "CCRMB", "RA", "L-RD", "SQUARE RADIANT MODIFIED", "RADIANT MODIFIED BRILLIANT", "RADIANT MODIFIED", "LR_BRILLIANT", "LONG RADIANT", "SQ.RADIANT", "RADIANT", "SQ RADIANT"],
    "SQ.EMERALD": ["SEM", "SQ.EMERALD", "S.EMERALD", "SQ.EM", "SE", "SQ-EM", "SQ.EMERALD"],
    BAGUETTE: ["BT", "RC", "RSC", "BUG", "BU", "BUGGETTE", "BAGUETTE"] ,
    BRIOLETTE: ["BCM", "BLT", "BRIOLETTE"],
    "CUS.BRILLIANT": ["CB", "CR", "CUS.BRILLIANT"],
  },
  Color: {
    D: ["D", "*D"],
    E: ["E"],
    F: ["F", "F YELLO"],
    G: ["G", "RARE WHITE G"],
    H: ["H", "H-", "WHITE (H)", "WHITE H"],
    I: ["I"],
    K: ["K", "K, Faint Brown"],
    L: ["L", "L, Faint Brown"],
    M: ["M", "TINTED COLOUR M"],
    N: ["N"],
    X: ["X", "XYZ"],
  },
  Clarity: {
    VVS1: ["VVS1", "VVS1-", "vvs1", "VVS 1"],
    VVS2: ["VVS2", "VVS2-", "vvs2", "VVS 2"],
    IF: ["IF", "LOUPE-CLEAN", "LC", "Internally Flawless"],
    SI1: ["SI1", "SI1-", "SI 1"],
    SI2: ["SI2", "SI2-", "SI 2"],
  },
  Cut: {
    EX: ["EXCELLENT", "Excellent", "EXC", "IDEAL", "Ideal", "Ex", "EX-2", "EX-1", "EX-3", "X1", "X2", "X3", "X4", "ex1", "ex2", "I", "Ex Ideal", "EX3"],
    VG: ["VERY GOOD", "V. GOOD", "VG", "VV", "VX", "V"],
    GD: ["GOOD", "Good", "GD", "G", "GV", "GX", "ex6", "G", "Gd"],
    FR: ["FAIR", "F", "FR", "FA", "F"],
    PR: ["P", "POOR", "PR", "PU"],
  },
  Polish: {
    EX: ["EXCELLENT", "Ex", "EXC", "IDEAL", "Ideal", "X1", "X2", "EX-2", "Ex Ideal"],
    VG: ["VERY GOOD", "VG", "V. GOOD"],
    GD: ["GOOD", "Good", "GD", "G", "Gd"],
    FR: ["FAIR", "FR", "F", "FA"],
    PR: ["P", "POOR", "PR"],
  },
  Symm: {
    EX: ["EXCELLENT", "Ex", "EXC", "IDEAL", "Ideal", "X1", "X2", "EX-2", "Ex Ideal"],
    VG: ["VERY GOOD", "VG", "V. GOOD"],
    GD: ["GOOD", "Good", "GD", "G", "Gd"],
    FR: ["FAIR", "FR", "F", "FA"],
    PR: ["P", "POOR", "PR"],
  },
  Fls: {
    NONE: ["N", "NO", "NIL", "FL0", "NN", "Non", "None"],
    FNT: ["FAINT", "FNT", "NEGLIGIBLE", "FL1", "FA"],
    MED: ["M", "MEDIUM", "MED", "FL2", "MD", "MB", "MEDIUM BLUE", "ME", "MD-BL", "MD-YL"],
    STG: ["STRONG", "STG", "ST", "FL3", "S", "STRONG BLUE", "STR", "STO", "ST-BL", "ST-WH", "ST-YL"],
    VST: ["VERY STRONG", "VST", "FL4", "VSTG", "VST-BL"],
    SLT: ["SL", "SLIGHT", "SLI", "SLT"],
    VSL: ["VERY SLIGHT", "VSLG", "VSLT", "VSL", "VS"],
  },
  Lab: {
    GIA: ["G.I.A", "GIA", "GA", "GIA"],
    IGI: ["IGI"],
    "NO-CERT": ["NONE", "NON", "NC", "NONCERT", "NCERT", "NON CERT"],
  },
};

const ColumnSynonyms = {
  Shape: ["Shape Name", "SHAPE#", "SHAP#", "Shp#", "Shape Desc", "Rap Shape", "rapShape", "Sh", "Sha", "SHAP", "SHAP.", "Shape", "Shape ??", "SHAPE.", "Shape_Code", "Shp", "Shp.", "Shp_Name"],
  Color: ["Colour", "Color (Long)", "Color C", "Col#", "CL.", "clours", "CLR", "Col", "Col.", "Color", "Color ??", "Color_Code", "ColorCode", "Colour", "Rap Color", "rapColor"],
  Clarity: ["Cl#", "Quality", "clar", "Cla#", "CAL", "Cal_Name", "CL", "Cl.", "Cla", "Cla.", "Clar", "Claratiy", "Clari", "Clarity", "Clarity ??", "Clarity_Code", "ClarityName", "CLERITY", "CLR", "Clr.", "Clrt", "CLRTY", "CTY", "Purity", "Qua", "Rap Clarity", "rapClarity"],
  Cut: ["Proportions", "CUTPROP", "CUT#", "Prop#", "Prop", "Ct", "CUT", "Cut Grade", "Cut Grade ??", "CUT.", "Cut_Code", "Cut_Grade", "CutGrade", "Final Cut", "Prop.", "PropCode"],
  Polish: ["POL#", "PL", "po", "POL", "POL.", "Polish", "Polish ??", "Polish_Code", "PolishName"],
  Symm: ["Symm", "SYM#", "SUM", "SYS", "Sy", "SYM", "Sym.", "SYMM", "Symmetry", "Symmetry ??", "Symmetry_Code", "Symmmetry", "SymName"],
  Fls: ["Flr", "FLRN", "FLUOR#", "Flour#", "FLRInt", "Fluore#", "Flor#", "Flo#", "FL", "FL.", "FLName", "FLO", "flor", "FLOR.", "Flore", "Floro", "FLORO.", "Florosence", "Flors", "Flou", "Flour", "Flour.", "Flourence", "FLOURESCENSE", "FLOURESENCE", "Flr", "flr_intensity", "Flrcnt", "FLRInt", "FlrIntens", "Fls", "FLS.", "FLU", "fluo", "Fluo Int", "Fluo.", "Fluor", "Fluor.", "Fluores..", "Fluorescence", "Fluorescence ??", "Fluorescence Intensity", "Fluorescence_Code", "Fluorescence_Intensity", "FluorescenceColor", "FluorescenceIntensity", "Fluorescense", "Fluorescent", "FLUORS", "Flur", "Fluro", "FLURO."],
  Lab: ["Cer", "Cert", "Cert By", "Cert From", "Cert Name", "CERT.", "CERTI", "CERTI_NAME", "CERTIFICATE", "CertName", "CR_Name", "Crt", "Lab", "Lab ??", "Lab Name", "Lab_Code", "report"],
};

// const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.add("A1:M1", true /*hasHeaders*/);
//     expensesTable.name = "ExpensesTable";

//     expensesTable.getHeaderRowRange().values = [["Shape", "Clarity", "Cut", "Polish", "Symm", "Fls", "Lab", "Weight", "Rate", "Disc"]];

//     expensesTable.rows.add(null /*add at the end*/, [
//       ["Round", "VVS1",	"Good",	"Good",	"Good",	"None",	"G.I.A",	"0.25",	"5000",	"-25"],
//       ["Ht", "VVS2",	"Ideal",	"Ideal",	"Ideal",	"Non",	"GIA",	"0.98",	"6000",	"-27"]
//       ["Em", "IF",	"Excellent",	"Excellent",	"Excellent",	"Green",	"IGI",	"1.25",	"15000", "30"]
//       ["Em", "SI1",	"Ex",	"Ex",	"Ex",	"Non",	"HRD",	"0.6",	"5800",	"-31"]
//       ["Round",	"SI2",	"Gd",	"Gd",	"Gd",	"None",	"NCERT",	"0.4",	"5500",	"-50"]
//     ]);

//     expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
//     expensesTable.getRange().format.autofitColumns();
//     expensesTable.getRange().format.autofitRows();
//     await context.sync();

// Shape: [Shape Name|SHAPE#|SHAP#|Shp#|Shape Desc|Rap Shape|rapShape|Sh|Sha|SHAP|SHAP.|Shape|Shape ??|SHAPE.|Shape_Code|Shp|Shp.|Shp_Name]
// Color: [Colour|Color (Long)|Color C|Col#|CL.|clours|CLR|Col|Col.|Color|Color ??|Color_Code|ColorCode|Colour|Rap Color|rapColor]
// Clarity: [Cl#|Quality|clar|Cla#|CAL|Cal_Name|CL|Cl.|Cla|Cla.|Clar|Claratiy|Clari|Clarity|Clarity ??|Clarity_Code|ClarityName|CLERITY|CLR|Clr.|Clrt|CLRTY|CTY|Purity|Qua|Rap Clarity|rapClarity]
// Cut: [Proportions|CUTPROP|CUT#|Prop#|Prop|Ct|CUT|Cut Grade|Cut Grade ??|CUT.|Cut_Code|Cut_Grade|CutGrade|Final Cut|Prop.|PropCode]
// Polish: [POL#|PL|po|POL|POL.|Polish|Polish ??|Polish_Code|PolishName]
// Symm: [Symm|SYM#|SUM|SYS|Sy|SYM|Sym.|SYMM|Symmetry|Symmetry ??|Symmetry_Code|Symmmetry|SymName]
// Fls: [Flr|FLRN|FLUOR#|Flour#|FLRInt|Fluore#|Flor#|Flo#|FL|FL.|FLName|FLO|flor|FLOR.|Flore|Floro|FLORO.|Florosence|Flors|Flou|Flour|Flour.|Flourence|FLOURESCENSE|FLOURESENCE|Flr|flr_intensity|Flrcnt|FLRInt|FlrIntens|Fls|FLS.|FLU|fluo|Fluo Int|Fluo.|Fluor|Fluor.|Fluores..|Fluorescence|Fluorescence ??|Fluorescence Intensity|Fluorescence_Code|Fluorescence_Intensity|FluorescenceColor|FluorescenceIntensity|Fluorescense|Fluorescent|FLUORS|Flur|Fluro|FLURO.]
// Lab: [Cer|Cert|Cert By|Cert From|Cert Name|CERT.|CERTI|CERTI_NAME|CERTIFICATE|CertName|CR_Name|Crt|Lab|Lab ??|Lab Name|Lab_Code|report]