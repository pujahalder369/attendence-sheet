import * as XLSX from "xlsx-js-style";

// =====================================================
// BORDER
// =====================================================

export const borderStyle = {
  top: {
    style: "thin",
    color: { rgb: "000000" },
  },
  bottom: {
    style: "thin",
    color: { rgb: "000000" },
  },
  left: {
    style: "thin",
    color: { rgb: "000000" },
  },
  right: {
    style: "thin",
    color: { rgb: "000000" },
  },
};

// =====================================================
// CENTER
// =====================================================

const centerAlignment = {
  horizontal: "center",
  vertical: "center",
  wrapText: true,
};

// =====================================================
// GENERAL BORDER
// =====================================================

export const applyBorderToAllCells = (worksheet) => {
  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;

    const cell = worksheet[cellAddress];

    if (!cell) return;

    cell.s = {
      ...(cell.s || {}),

      border: {
        ...borderStyle,
      },

      alignment: {
        ...(cell.s?.alignment || {}),
        ...centerAlignment,
      },
    };
  });
};

// =====================================================
// EMPLOYEE HEADER
// =====================================================

export const applyEmployeeHeaderStyle = (worksheet) => {
  for (let row = 0; row <= 3; row++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: 0,
    });

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
      font: {
        bold: true,
        sz: 11,
        color: {
          rgb: "000000",
        },
      },

      alignment: {
        horizontal: "left",
        vertical: "center",
      },

      border: {
        ...borderStyle,
      },
    };
  }
};

// =====================================================
// SUMMARY HEADER
// =====================================================

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

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
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
          rgb: "404040",
        },
      },

      alignment: {
        ...centerAlignment,
      },

      border: {
        ...borderStyle,
      },
    };
  }
};

// =====================================================
// SUMMARY DATA
// =====================================================

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

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
      font: {
        bold: true,
        sz: 10,
        color: {
          rgb: "000000",
        },
      },

      alignment: {
        ...centerAlignment,
      },

      border: {
        ...borderStyle,
      },
    };
  }
};

// =====================================================
// DAY HEADER
// =====================================================

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

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
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

      alignment: {
        ...centerAlignment,
      },

      border: {
        ...borderStyle,
      },
    };
  }
};

// =====================================================
// LABEL STYLE
// =====================================================

export const applyLabelStyle = (
  worksheet,
  rows
) => {
  rows.forEach((row) => {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: 0,
    });

    const cell = worksheet[cellAddress];

    if (!cell) return;

    cell.s = {
      font: {
        bold: true,
        sz: 10,
        color: {
          rgb: "000000",
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
      },

      border: {
        ...borderStyle,
      },
    };
  });
};

// =====================================================
// STATUS COLORS
// =====================================================

const statusColors = {
  P: "C6EFCE",   // Present
  A: "FFC7CE",   // Absent
  L: "FFEB9C",   // Leave
  HD: "F4B183",  // Half Day
  HO: "9DC3E6",  // Holiday
  WO: "D9D9D9",  // Week Off
};

// =====================================================
// STATUS STYLE
// =====================================================

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

    const background =
      statusColors[status] || "FFFFFF";

    cell.s = {
      font: {
        bold: true,
        sz: 10,
        color: {
          rgb: "000000",
        },
      },

      fill: {
        patternType: "solid",
        fgColor: {
          rgb: background,
        },
      },

      alignment: {
        ...centerAlignment,
      },

      border: {
        ...borderStyle,
      },
    };
  }
};