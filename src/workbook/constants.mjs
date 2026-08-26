export const SHEET_SETTINGS = "00_\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438";
export const SHEET_QUARTER = "01_\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043a\u0432\u0430\u0440\u0442\u0430\u043b\u0430";
export const SHEET_CAPACITY = "02_Capacity";
export const SHEET_ESTIMATES = "03_\u041e\u0446\u0435\u043d\u043a\u0430 \u0437\u0430\u0434\u0430\u0447";
export const SHEET_PLAN = "04_\u041a\u0432\u0430\u0440\u0442\u0430\u043b\u044c\u043d\u044b\u0439 \u043f\u043b\u0430\u043d";
export const SHEET_EXPRESS_TEMPLATE = "100_\u0428\u0430\u0431\u043b\u043e\u043d \u044d\u043a\u0441\u043f\u0440\u0435\u0441\u0441 \u043e\u0446\u0435\u043d\u043a\u0438";
export const SHEET_REFS = "99_\u0421\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a\u0438";
export const EXCEL_DATE_FORMAT = "dd-mm-yyyy";

export function excelSheetName(name) {
  return Array.from(name).slice(0, 31).join("");
}

export function sheetRef(name) {
  return `'${excelSheetName(name).replace(/'/g, "''")}'`;
}
