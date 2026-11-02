function getPickingConstants() {
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
    customFields
  });
}

function onOpen() {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Tableau")
  .addItem("Update Tableau Data", updatePickingTableauData.name)
  .addToUi();
}

function updatePickingTableauData() {
  const C = getConstants();
  const RAW_DATA = getRawPickingTableauData(C);
  const FIXED_DATA = reformatPickingTableauData(C, RAW_DATA);
  const NEW_DATA = filterForNewData(C, FIXED_DATA);
  pastePickingTableauData(C, NEW_DATA);
}
