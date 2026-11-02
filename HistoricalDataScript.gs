function getConstants() {
  const customFields = [
    {
      name: "Logical Date",
      func: {
        paramFields: ["SQ First Processed (ET)", "Time to Process"],
        callback: getLogicalDatePicking
      }
    }
  ];

  return Object.freeze({
    S: SpreadsheetApp.getActive(),
    tableauId: "5cfa71e7-1f06-46d8-a40b-475901640afd",
    measureHeader: "Measure Names",
    measureField: "Measure Values",
    getImportSheet: function() {
      const id = 0;
      return this.S.getSheetById(id);
    },
    getHistoricalDataSheet: function() {
      const id = 683115303;
      return this.S.getSheetById(id);
    },
    getHistoricalIds: function() {
      const SS = this.getHistoricalDataSheet();
      const dataPoints = SS.getLastRow() - 1;
      return new Set(dataPoints > 0 ? SS.getRange(2, 1, dataPoints, 1).getValues().flat() : "");
    },
    customFields
  });
}


/**
 * Calls the Tableau API and grabs the unedited report data.
 * @param {Object} c - The C constants object. Creates it if not passed through.
 * @returns {Array, Array} headers, rows - The 1D headers array, and the 2D report data.
 * Headers have been included from legacy, but are no longer used.
 */
function getRawPickingTableauData(C) {
  C = C ?? getConstants();
  C.S.toast("Getting raw Tableau data...");

  const api = tableau.APIRequest().signin("mbaxter@ebay.com");
  let data = api.getViewData(C.tableauId).execute();

  data = Utilities.parseCsv(data);
  const headers = data.shift();
  return {headers, rows: data};
}

/**
 * Reformats the raw Tableau data from Measure Names and Values to proper columns.
 * @param {Object} C - The C constants object. Creates it if not passed through.
 * @param {Object} data - {headers: Array, rows: Array} The unedited data from Tableau, with headers and rows.
 * @returns {Array, Array} headers, rows - The reformatted arrays of data.
 */
function reformatPickingTableauData(C, data) {
  C = C ?? getConstants();
  const measures = determineMeasuresPicking(C, data);

  if (data.rows.length % measures.length !== 0) return [["Insufficient data points compared to measure names."]];

  const fieldCount = data.headers.length - 1;

  const correctedHeaders = formatHeadersPicking([
    ...data.headers.slice(1, fieldCount), 
    ...measures,
    ...C.customFields.map(f => f.name)
  ]);

  const output = [];
  for (let i = 0; i < data.rows.length; i += measures.length) {
    const row = data.rows[i].slice(1, fieldCount);
    if (row[0] === "") continue;
    for (let j = 0; j < measures.length; j++) {
      row.push(data.rows[i + j][fieldCount]);
    }
    C.customFields.forEach(field => {
      const {paramFields, callback} = field.func;
      const params = formatHeadersPicking(paramFields).map(f => row[correctedHeaders.indexOf(f)]);
      row.push(callback(...params));
    });
    output.push(row);
  }

  return {headers: correctedHeaders, rows: output};
}

function filterForNewData(C, formattedData) {
  C = C ?? getConstants();
  const historicalIds = C.getHistoricalIds();
  const newData = [];
  for (const row of formattedData.rows) {
    const [name, sq] = [row[0], row[5]];
    const id = `${name}|${sq}`;
    if (!historicalIds.has(id)) {
      newData.push([id, ...row]);
    }
  }
  return newData;
}

/**
 * Takes the passed data and pastes it into the designated import sheet.
 * @param {Object} C - the C constants object. Creates it if not passed through.
 * @param {Array} data - The data to be pasted into the designated import sheet.
 */
function pastePickingTableauData(C, data) {
  C = C ?? getConstants();
  C.S.toast("Pasting Tableau data into DataImport sheet.");
  const SS = C.getHistoricalDataSheet();
  const {dLen, dWdth} = determineDataSizePicking(data);
  if (!dLen || !dWdth) return;

  const nextRow = SS.getLastRow() + 1;
  SS.getRange(nextRow, 1, dLen, dWdth).setValues(data);
}

/**
 * Determines the dimensions of the data array (length and "width").
 * @param {Array} data - The unformatted Tableau data, excluding headers.
 * @returns {number, number} - the length of the entire data array, and the length of each row (should be uniform).
 */
function determineDataSizePicking(data) {
  let dLen, dWdth;
  if (!data || !Array.isArray(data) || data.length < 1) {
    dLen = 0, dWdth = 0;
  }
  else {
    dLen = data.length, dWdth = data[0].length;
  }
  return {dLen, dWdth}; // Might want to improve the safety of this function some day and have it confirm all rows are equal length.
}

/**
 * Collects the unique measures in the Tableau report and returns them as an array.
 * @param {Object} c - the C constants object. Creates it if not passed through.
 * @param {Object} fullData - {headers, rows} The unformatted Tableau data, output of getRawTableauData.
 * @returns {Array} - An array of the unique measures within the Tableau report.
 */
function determineMeasuresPicking(c, fullData) {
  if (!fullData || !fullData.headers || !fullData.rows || !Array.isArray(fullData.headers) || !Array.isArray(fullData.rows)) return [];

  const C = c ?? getPickingConstants();
  const mIdx = fullData.headers.indexOf(C.measureHeader);
  if (mIdx === -1) return [];

  let measures = new Set();
  for (const row of fullData.rows) {
    const measure = row[mIdx];
    if (!measures.has(measure)) measures.add(measure);
    else break;
  }

  measures = [...measures];
  return measures;
}

function getLogicalDatePicking(startDate, timeToProcess) {
  if (!startDate || !timeToProcess) return "";
  const date = new Date(startDate);
  const durationDate = new Date(timeToProcess);

  const [h, m, s] = [durationDate.getHours(), durationDate.getMinutes(), durationDate.getSeconds()];
  
  /*
  This section isn't working as intended because Tableau returns a date object as a string, which breaks the string check.
  let h, m, s;
  if (timeToProcess instanceof Date) {
    h = timeToProcess.getHours();
    m = timeToProcess.getMinutes();
    s = timeToProcess.getSeconds();
  }
  else if (typeof timeToProcess === "string") {
    [h, m, s] = timeToProcess.split(":").map(n => Number(n));
  }
  else {
    return "";
  }
  */

  date.setHours(date.getHours() + h, date.getMinutes() + m, date.getSeconds() + s);
  return new Date(date).toLocaleDateString();
}

function formatHeadersPicking(headers) {
  const formatRules = (h) => {
    let h2 = String(h).trim().toUpperCase();
    return h2.replace(/ /g, "_").replace(/\./g, "");
  }
  return headers.map(formatRules);
}
