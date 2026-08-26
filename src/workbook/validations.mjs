export function addTable(sheet, range, name) {
  const table = sheet.tables.add(range, true, name);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  return table;
}

export function addListValidation(range, values) {
  range.dataValidation = {
    rule: { type: "list", values },
    prompt: { showPrompt: true },
    errorAlert: { showAlert: true },
  };
}

export function addWholeValidation(range, min, max) {
  range.dataValidation = {
    rule: { type: "whole", operator: "between", formula1: min, formula2: max },
    errorAlert: { showAlert: true },
  };
}

export function addDecimalValidation(range, min, max) {
  range.dataValidation = {
    rule: { type: "decimal", operator: "between", formula1: min, formula2: max },
    errorAlert: { showAlert: true },
  };
}

export function addNonNegativeValidation(range) {
  range.dataValidation = {
    rule: { type: "decimal", operator: "between", formula1: 0, formula2: 999999 },
    ignoreBlank: true,
    errorAlert: { showAlert: true },
  };
}
