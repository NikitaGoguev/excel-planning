export const DESIGN_FONT_NAME = "Calibri";
export const DESIGN_FONT_SIZE = 11;

export const colors = Object.freeze({
  title: "#1F4E79",
  section: "#9DC3E6",
  header: "#D9EAF7",
  input: "#FFF2CC",
  calculated: "#E2F0D9",
  technical: "#F3F3F3",
  border: "#808080",
  text: "#1F1F1F",
  mutedText: "#808080",
  white: "#FFFFFF",
});

export function fontStyle(overrides = {}) {
  return { name: DESIGN_FONT_NAME, size: DESIGN_FONT_SIZE, ...overrides };
}

function borderedFormat(fill, font, options = {}) {
  return {
    fill,
    font,
    verticalAlignment: options.verticalAlignment ?? "center",
    horizontalAlignment: options.horizontalAlignment,
    wrapText: options.wrapText ?? true,
    borders: {
      top: { style: "continuous", color: options.borderColor ?? colors.border },
      bottom: { style: "continuous", color: options.borderColor ?? colors.border },
      left: { style: "continuous", color: options.borderColor ?? colors.border },
      right: { style: "continuous", color: options.borderColor ?? colors.border },
    },
  };
}

export function applyTitle(range) {
  range.format = {
    fill: colors.title,
    font: fontStyle({ bold: true, color: colors.white }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

export function applyHeader(range) {
  range.format = borderedFormat(colors.header, fontStyle({ bold: true, color: colors.text }), {
    horizontalAlignment: "center",
  });
}

export function applyInput(range) {
  range.format = borderedFormat(colors.input, fontStyle({ color: colors.text }));
}

export function applyCalculated(range) {
  range.format = borderedFormat(colors.calculated, fontStyle({ color: colors.text }));
}

export function applyPlain(range) {
  range.format = borderedFormat(undefined, fontStyle({ color: colors.text }), { verticalAlignment: "top" });
}

export function applyAction(range) {
  range.format = borderedFormat(colors.header, fontStyle({ bold: true, color: colors.title }), {
    horizontalAlignment: "center",
  });
}

export function setWidths(sheet, widths) {
  for (const [column, widthPx] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidthPx = widthPx;
  }
}

export function applyCapacityRangeStyle(range, fill, horizontalAlignment = "left", bold = false) {
  range.format = borderedFormat(fill, fontStyle({ bold, color: colors.text }), {
    horizontalAlignment,
    borderColor: "#000000",
  });
}
