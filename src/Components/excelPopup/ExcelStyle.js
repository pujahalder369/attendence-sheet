import * as XLSX from "xlsx-js-style";

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

export const centerAlignment = {
  horizontal: "center",
  vertical: "center",
  wrapText: true,
};

export const leftAlignment = {
  horizontal: "left",
  vertical: "center",
  wrapText: true,
};


// EMPLOYEE INFORMATION
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

  alignment: leftAlignment,

  border: borderStyle,
};


// SUMMARY HEADER
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

  alignment: centerAlignment,

  border: borderStyle,
};


// SUMMARY DATA
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

  alignment: centerAlignment,

  border: borderStyle,
};


// DAY HEADER
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

  alignment: centerAlignment,

  border: borderStyle,
};


// LABEL
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

  alignment: leftAlignment,

  border: borderStyle,
};


// STATUS COLORS
export const statusConfig = {
  P: {
    background: "C6EFCE",
    font: "006100",
  },

  A: {
    background: "FFC7CE",
    font: "9C0006",
  },

  L: {
    background: "FFEB9C",
    font: "9C6500",
  },

  HD: {
    background: "F4B084",
    font: "843C0C",
  },

  HO: {
    background: "9DC3E6",
    font: "1F4E78",
  },

  WO: {
    background: "D9E1F2",
    font: "404040",
  },
};


// APPLY BORDER
export const applyBorderToAllCells = (worksheet) => {
  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;

    const cell = worksheet[cellAddress];

    cell.s = {
      ...(cell.s || {}),

      alignment: {
        ...centerAlignment,
        ...(cell.s?.alignment || {}),
      },

      border: {
        ...borderStyle,
        ...(cell.s?.border || {}),
      },
    };
  });
};


// APPLY EMPLOYEE HEADER
export const applyEmployeeHeaderStyle = (worksheet) => {
  ["A1", "A2", "A3", "A4"].forEach((cellAddress) => {
    const cell = worksheet[cellAddress];

    if (!cell) return;

    cell.s = {
      ...employeeHeaderStyle,
    };
  });
};


// GENERIC ROW STYLE
const applyRowStyle = (
  worksheet,
  row,
  totalColumns,
  style
) => {
  for (let col = 0; col < totalColumns; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: row,
      c: col,
    });

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
      ...style,
    };
  }
};


// SUMMARY HEADER
export const applyHeaderStyle = (
  worksheet,
  row,
  totalColumns
) => {
  applyRowStyle(
    worksheet,
    row,
    totalColumns,
    headerStyle
  );
};


// SUMMARY DATA
export const applySummaryDataStyle = (
  worksheet,
  row,
  totalColumns
) => {
  applyRowStyle(
    worksheet,
    row,
    totalColumns,
    summaryDataStyle
  );
};

// DAY HEADER
export const applyDayHeaderStyle = (
  worksheet,
  row,
  totalColumns
) => {
  applyRowStyle(
    worksheet,
    row,
    totalColumns,
    dayHeaderStyle
  );
};

// LABEL COLUMN
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
      ...labelStyle,
    };
  });
};

// STATUS STYLE

export const applyStatusStyle = (
  worksheet,
  statusRow,
  totalDays
) => {
  for (let col = 1; col <= totalDays; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: statusRow,
      c: col,
    });

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    const config = statusConfig[cell.v];

    // Default style
    if (!config) {
      cell.s = {
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
            rgb: "FFFFFF",
          },
        },

        alignment: centerAlignment,

        border: borderStyle,
      };

      continue;
    }

    cell.s = {
      font: {
        bold: true,
        color: {
          rgb: config.font,
        },
        sz: 10,
      },

      fill: {
        patternType: "solid",
        fgColor: {
          rgb: config.background,
        },
      },

      alignment: centerAlignment,

      border: borderStyle,
    };
  }
};