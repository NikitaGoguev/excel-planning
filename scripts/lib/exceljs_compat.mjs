import ExcelJS from "exceljs";

function columnNumber(label) {
  return Array.from(label.toUpperCase()).reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function columnLabel(number) {
  let value = number;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function parseReference(reference) {
  const [startText, endText = startText] = reference.split(":");
  const parsePart = (part) => {
    const match = part.match(/^([A-Z]+)(\d+)?$/i);
    if (!match) throw new Error(`Unsupported A1 reference: ${reference}`);
    return { column: columnNumber(match[1]), row: match[2] ? Number(match[2]) : null };
  };
  const start = parsePart(startText);
  const end = parsePart(endText);
  if ((start.row === null) !== (end.row === null)) throw new Error(`Mixed row/column reference: ${reference}`);
  return {
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
    startRow: start.row === null ? null : Math.min(start.row, end.row),
    endRow: start.row === null ? null : Math.max(start.row, end.row),
  };
}

function argb(value) {
  if (!value) return undefined;
  const hex = String(value).replace(/^#/, "").toUpperCase();
  return hex.length === 8 ? hex : `FF${hex}`;
}

function normalizeFill(fill) {
  if (!fill) return undefined;
  return { type: "pattern", pattern: "solid", fgColor: { argb: argb(fill) } };
}

function normalizeFont(font) {
  if (!font) return undefined;
  const normalized = { ...font };
  if (font.color) normalized.color = { argb: argb(font.color) };
  return normalized;
}

function normalizeBorder(border) {
  if (!border) return undefined;
  const normalized = {};
  for (const side of ["top", "bottom", "left", "right", "diagonal"]) {
    if (!border[side]) continue;
    normalized[side] = {
      style: border[side].style === "continuous" ? "thin" : border[side].style,
      color: border[side].color ? { argb: argb(border[side].color) } : undefined,
    };
  }
  return normalized;
}

function clone(value) {
  return value === undefined ? undefined : structuredClone(value);
}

class RangeAdapter {
  constructor(sheet, reference) {
    this.sheet = sheet;
    this.reference = reference;
    this.bounds = parseReference(reference);
  }

  forEachCell(callback) {
    const { startColumn, endColumn, startRow, endRow } = this.bounds;
    if (startRow === null) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        this.sheet.raw.getColumn(column).eachCell({ includeEmpty: true }, (cell, row) => callback(cell, row, column));
      }
      return;
    }
    for (let row = startRow; row <= endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        callback(this.sheet.raw.getCell(row, column), row, column);
      }
    }
  }

  set values(matrix) {
    const { startColumn, startRow } = this.bounds;
    if (startRow === null) throw new Error(`Cannot assign values to a column-only range: ${this.reference}`);
    matrix.forEach((rowValues, rowOffset) => {
      rowValues.forEach((value, columnOffset) => {
        this.sheet.raw.getCell(startRow + rowOffset, startColumn + columnOffset).value = value;
      });
    });
  }

  set formulas(matrix) {
    const { startColumn, startRow } = this.bounds;
    if (startRow === null) throw new Error(`Cannot assign formulas to a column-only range: ${this.reference}`);
    matrix.forEach((rowValues, rowOffset) => {
      rowValues.forEach((formula, columnOffset) => {
        const cell = this.sheet.raw.getCell(startRow + rowOffset, startColumn + columnOffset);
        cell.value = formula === null || formula === ""
          ? formula
          : { formula: String(formula).replace(/^=/, "") };
      });
    });
  }

  set dataValidation(config) {
    const rule = config.rule ?? {};
    const validation = {
      type: rule.type,
      operator: rule.operator,
      allowBlank: config.ignoreBlank ?? true,
      showInputMessage: config.prompt?.showPrompt ?? false,
      showErrorMessage: config.errorAlert?.showAlert ?? true,
    };
    if (rule.type === "list") {
      validation.formulae = [`"${(rule.values ?? []).join(",")}"`];
    } else {
      validation.formulae = [rule.formula1, rule.formula2].filter((value) => value !== undefined);
    }
    this.forEachCell((cell) => {
      cell.dataValidation = clone(validation);
    });
  }

  set format(format) {
    this.applyFormat(format);
  }

  get format() {
    return new Proxy({}, {
      set: (_target, property, value) => {
        this.applyFormat({ [property]: value });
        return true;
      },
    });
  }

  applyFormat(format) {
    if (format.columnWidthPx !== undefined) {
      const width = Math.max(1, (Number(format.columnWidthPx) - 5) / 7);
      for (let column = this.bounds.startColumn; column <= this.bounds.endColumn; column += 1) {
        this.sheet.raw.getColumn(column).width = width;
      }
    }
    if (format.columnHidden !== undefined) {
      for (let column = this.bounds.startColumn; column <= this.bounds.endColumn; column += 1) {
        this.sheet.raw.getColumn(column).hidden = Boolean(format.columnHidden);
      }
    }

    const hasCellStyle = ["fill", "font", "borders", "horizontalAlignment", "verticalAlignment", "wrapText"]
      .some((key) => Object.hasOwn(format, key));
    if (!hasCellStyle) return;

    this.forEachCell((cell) => {
      if (Object.hasOwn(format, "fill")) cell.fill = normalizeFill(format.fill);
      if (Object.hasOwn(format, "font")) cell.font = normalizeFont(format.font);
      if (Object.hasOwn(format, "borders")) cell.border = normalizeBorder(format.borders);
      const alignment = { ...(cell.alignment ?? {}) };
      if (Object.hasOwn(format, "horizontalAlignment")) alignment.horizontal = format.horizontalAlignment;
      if (Object.hasOwn(format, "verticalAlignment")) alignment.vertical = format.verticalAlignment;
      if (Object.hasOwn(format, "wrapText")) alignment.wrapText = Boolean(format.wrapText);
      cell.alignment = alignment;
    });
  }

  merge(acrossRows = false) {
    const { startColumn, endColumn, startRow, endRow } = this.bounds;
    if (startRow === null) throw new Error(`Cannot merge a column-only range: ${this.reference}`);
    if (acrossRows) {
      for (let row = startRow; row <= endRow; row += 1) {
        this.sheet.raw.mergeCells(row, startColumn, row, endColumn);
      }
      return;
    }
    this.sheet.raw.mergeCells(startRow, startColumn, endRow, endColumn);
  }

  setNumberFormat(numberFormat) {
    this.forEachCell((cell) => {
      cell.numFmt = numberFormat;
    });
  }

  fillDown() {
    const { startColumn, startRow } = this.bounds;
    if (startRow === null) throw new Error(`Cannot fill a column-only range: ${this.reference}`);
    const source = this.sheet.raw.getCell(startRow, startColumn).value;
    if (!source || typeof source !== "object" || !source.formula) {
      throw new Error(`Top cell does not contain a formula: ${this.reference}`);
    }
    this.sheet.raw.fillFormula(this.reference, source.formula);
  }
}

class TableAdapter {
  constructor(table) {
    this.raw = table;
  }

  set style(value) {
    this.raw.style = { theme: value, showRowStripes: false, showColumnStripes: false };
  }

  set showFilterButton(_value) {
    // ExcelJS emits the table autoFilter. Per-column hidden buttons are patched in OOXML later.
  }
}

class SheetAdapter {
  constructor(raw) {
    this.raw = raw;
    this.freezePanes = {
      freezeRows: (rows) => {
        const current = this.raw.views?.[0] ?? {};
        this.raw.views = [{ ...current, state: "frozen", ySplit: rows, activeCell: `A${rows + 1}` }];
      },
    };
    this.tables = {
      add: (reference, headerRow, name) => {
        const bounds = parseReference(reference);
        if (!headerRow || bounds.startRow === null) throw new Error(`Unsupported table range: ${reference}`);
        const headers = [];
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          headers.push(String(this.raw.getCell(bounds.startRow, column).value ?? ""));
        }
        const rows = [];
        for (let row = bounds.startRow + 1; row <= bounds.endRow; row += 1) {
          const values = [];
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            values.push(clone(this.raw.getCell(row, column).value));
          }
          rows.push(values);
        }
        const table = this.raw.addTable({
          name,
          ref: `${columnLabel(bounds.startColumn)}${bounds.startRow}`,
          headerRow: true,
          totalsRow: false,
          style: { theme: "TableStyleMedium2", showRowStripes: false, showColumnStripes: false },
          columns: headers.map((header) => ({ name: header })),
          rows,
        });
        return new TableAdapter(table);
      },
    };
  }

  getRange(reference) {
    return new RangeAdapter(this, reference);
  }

  set showGridLines(value) {
    const current = this.raw.views?.[0] ?? {};
    this.raw.views = [{ ...current, showGridLines: Boolean(value) }];
  }
}

export class WorkbookAdapter {
  constructor() {
    this.raw = new ExcelJS.Workbook();
    this.raw.creator = "QuarterPlan Excel";
    this.raw.lastModifiedBy = "QuarterPlan Excel";
    this.raw.calcProperties.fullCalcOnLoad = true;
    this.raw.calcProperties.forceFullCalc = true;
    this.worksheets = {
      add: (name) => new SheetAdapter(this.raw.addWorksheet(name)),
    };
  }

  static create() {
    return new WorkbookAdapter();
  }
}

export async function appendWorksheetFromFile(targetWorkbook, sourcePath, targetName) {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourcePath);
  const source = sourceWorkbook.worksheets[0];
  if (!source) throw new Error(`Workbook has no worksheets: ${sourcePath}`);

  const target = targetWorkbook.addWorksheet(targetName, {
    properties: clone(source.properties),
    pageSetup: clone(source.pageSetup),
    views: clone(source.views),
  });
  target.headerFooter = clone(source.headerFooter);

  for (let column = 1; column <= source.columnCount; column += 1) {
    const sourceColumn = source.getColumn(column);
    const targetColumn = target.getColumn(column);
    if (sourceColumn.width !== undefined) targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
  }

  source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    if (sourceRow.height !== undefined) targetRow.height = sourceRow.height;
    targetRow.hidden = sourceRow.hidden;
    targetRow.outlineLevel = sourceRow.outlineLevel;
    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, columnNumberValue) => {
      const targetCell = targetRow.getCell(columnNumberValue);
      targetCell.value = clone(sourceCell.value);
      targetCell.style = clone(sourceCell.style);
      targetCell.dataValidation = clone(sourceCell.dataValidation);
      if (sourceCell.note !== undefined) targetCell.note = clone(sourceCell.note);
    });
  });

  for (const merge of source.model.merges ?? []) target.mergeCells(merge);
  return target;
}

export { ExcelJS };
