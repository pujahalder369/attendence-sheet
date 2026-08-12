import * as XLSX from "xlsx-js-style";

// =========================
// BORDER
// =========================

export const borderStyle = {
  top: {
    style: "thin",
    color: { rgb: "B7B7B7" },
  },
  bottom: {
    style: "thin",
    color: { rgb: "B7B7B7" },
  },
  left: {
    style: "thin",
    color: { rgb: "B7B7B7" },
  },
  right: {
    style: "thin",
    color: { rgb: "B7B7B7" },
  },
};

// =========================
// CENTER ALIGNMENT
// =========================

export const centerStyle = {
  horizontal: "center",
  vertical: "center",
  wrapText: true,
};

// =========================
// MAIN TITLE
// =========================

export const titleStyle = {
  font: {
    bold: true,
    color: { rgb: "FFFFFF" },
    sz: 14,
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "1F4E78",
    },
  },

  alignment: {
    horizontal: "center",
    vertical: "center",
  },

  border: borderStyle,
};

// =========================
// EMPLOYEE INFORMATION
// =========================

export const employeeHeaderStyle = {
  font: {
    bold: true,
    sz: 11,
    color: {
      rgb: "222222",
    },
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "D9EAF7",
    },
  },

  alignment: {
    horizontal: "left",
    vertical: "center",
    wrapText: true,
  },

  border: borderStyle,
};

// =========================
// DARK HEADER
// =========================

export const headerStyle = {
  font: {
    bold: true,
    color: {
      rgb: "FFFFFF",
    },
    sz: 11,
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "404040",
    },
  },

  alignment: centerStyle,

  border: borderStyle,
};

// =========================
// SUMMARY DATA
// =========================

export const summaryDataStyle = {
  font: {
    bold: true,
    color: {
      rgb: "000000",
    },
    sz: 10,
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "F3F6F9",
    },
  },

  alignment: centerStyle,

  border: borderStyle,
};

// =========================
// DAY HEADER
// =========================

export const dayHeaderStyle = {
  font: {
    bold: true,
    color: {
      rgb: "FFFFFF",
    },
    sz: 10,
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "5B9BD5",
    },
  },

  alignment: centerStyle,

  border: borderStyle,
};

// =========================
// LABEL COLUMN
// =========================

export const labelStyle = {
  font: {
    bold: true,
    color: {
      rgb: "222222",
    },
    sz: 10,
  },

  fill: {
    patternType: "solid",
    fgColor: {
      rgb: "EAF2F8",
    },
  },

  alignment: {
    horizontal: "left",
    vertical: "center",
    wrapText: true,
  },

  border: borderStyle,
};

// =========================
// STATUS COLORS
// =========================

export const statusColors = {
  P: "C6EFCE",
  A: "FFC7CE",
  L: "FFEB9C",
  HD: "F4B084",
  HO: "9DC3E6",
  WO: "D9E1F2",
};

// =========================
// STATUS FONT COLORS
// =========================

export const statusFontColors = {
  P: "006100",
  A: "9C0006",
  L: "9C6500",
  HD: "843C0C",
  HO: "1F4E78",
  WO: "404040",
};

// =========================
// APPLY BORDER
// =========================

export const applyBorderToAllCells = (worksheet) => {
  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;

    const oldStyle = worksheet[cellAddress].s || {};

    worksheet[cellAddress].s = {
      ...oldStyle,

      alignment: {
        ...centerStyle,
      },

      border: {
        ...borderStyle,
      },
    };
  });
};

// =========================
// EMPLOYEE HEADER STYLE
// =========================

export const applyEmployeeHeaderStyle = (worksheet) => {
  ["A1", "A2", "A3", "A4"].forEach((cellAddress) => {
    if (!worksheet[cellAddress]) return;

    worksheet[cellAddress].s = {
      ...employeeHeaderStyle,
    };
  });
};

// =========================
// TITLE STYLE
// =========================

export const applyTitleStyle = (worksheet) => {
  if (!worksheet.A1) return;

  worksheet.A1.s = {
    ...titleStyle,
  };
};

// =========================
// HEADER STYLE
// =========================

export const applyHeaderStyle = (
  worksheet,
  row,
  totalColumns
) => {
  for (let col = 0; col < totalColumns; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: col,
    });

    if (!worksheet[cellAddress]) continue;

    worksheet[cellAddress].s = {
      ...headerStyle,
    };
  }
};

// =========================
// SUMMARY DATA STYLE
// =========================

export const applySummaryDataStyle = (
  worksheet,
  row,
  totalColumns
) => {
  for (let col = 0; col < totalColumns; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: col,
    });

    if (!worksheet[cellAddress]) continue;

    worksheet[cellAddress].s = {
      ...summaryDataStyle,
    };
  }
};

// =========================
// DAY HEADER STYLE
// =========================

export const applyDayHeaderStyle = (
  worksheet,
  row,
  totalColumns
) => {
  for (let col = 0; col < totalColumns; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: col,
    });

    if (!worksheet[cellAddress]) continue;

    worksheet[cellAddress].s = {
      ...dayHeaderStyle,
    };
  }
};

// =========================
// LABEL STYLE
// =========================

export const applyLabelStyle = (
  worksheet,
  rows
) => {
  rows.forEach((row) => {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: 0,
    });

    if (!worksheet[cellAddress]) return;

    worksheet[cellAddress].s = {
      ...labelStyle,
    };
  });
};

// =========================
// STATUS STYLE
// =========================

export const applyStatusStyle = (
  worksheet,
  statusRow
) => {
  for (let col = 1; col <= 31; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: statusRow,
      c: col,
    });

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    const status = cell.v;

    const bgColor =
      statusColors[status] || "FFFFFF";

    const fontColor =
      statusFontColors[status] || "000000";

    cell.s = {
      font: {
        bold: true,
        color: {
          rgb: fontColor,
        },
        sz: 10,
      },

      fill: {
        patternType: "solid",
        fgColor: {
          rgb: bgColor,
        },
      },

      alignment: {
        horizontal: "center",
        vertical: "center",
      },

      border: {
        ...borderStyle,
      },
    };
  }
};