import * as XLSX from "xlsx-js-style";

export const COLORS = {
  header: "17365D",
  dayHeader: "2F75B5",
  white: "FFFFFF",
  black: "000000",
  border: "D9E1F2",
  employeeEven: "F8FAFC",
  employeeOdd: "EEF4FB",
  present: "C6EFCE",
  absent: "FFC7CE",
  leave: "FFEB9C",
  halfDay: "F4B183",
  holiday: "9DC3E6",
  weekOff: "D9D9D9",
};


// BORDER
export const borderStyle = {
  top: {
    style: "thin",
    color: {
      rgb: COLORS.border,
    },
  },

  bottom: {
    style: "thin",
    color: {
      rgb: COLORS.border,
    },
  },

  left: {
    style: "thin",
    color: {
      rgb: COLORS.border,
    },
  },

  right: {
    style: "thin",
    color: {
      rgb: COLORS.border,
    },
  },
};


// CENTER ALIGNMENT
export const centerAlignment = {
  horizontal: "center",
  vertical: "center",
  wrapText: true,
};


// GENERAL STYLE
export const applyGeneralStyle = (worksheet, totalRows, totalColumns) => {
  for (let row = 0; row < totalRows; row++) {
    for (let col = 0; col < totalColumns; col++) {
      const cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: col,
      });

      const cell = worksheet[cellAddress];
      if (!cell) continue;
      cell.s = {
        font: {
          sz: 10,
          color: {
            rgb: COLORS.black,
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
  }
};


// HEADER STYLE
export const applyHeaderStyle = (worksheet, totalColumns) => {
  for (let col = 0; col < totalColumns; col++) {
    const cellAddress = XLSX.utils.encode_cell({
      r: 0,
      c: col,
    });

    const cell = worksheet[cellAddress];

    if (!cell) continue;

    cell.s = {
      font: {
        bold: true,
        sz: 10,
        color: {
          rgb: COLORS.white,
        },
      },

      fill: {
        patternType: "solid",
        fgColor: {
          rgb: col < 10 ? COLORS.header : COLORS.dayHeader,
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


// EMPLOYEE STYLE
export const applyEmployeeStyle = (worksheet, totalRows) => {
  for (let row = 1; row < totalRows; row++) {
    for (let col = 0; col < 4; col++) {
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
            rgb: COLORS.black,
          },
        },

        fill: {
          patternType: "solid",

          fgColor: {
            rgb: row % 2 === 0 ? COLORS.employeeEven : COLORS.employeeOdd,
          },
        },

        alignment: {
          horizontal: col === 0 ? "center" : "left",

          vertical: "center",

          wrapText: true,
        },

        border: {
          ...borderStyle,
        },
      };
    }
  }
};


// SUMMARY STYLE
export const applySummaryStyle = (worksheet, totalRows) => {
  const summaryColors = {
    4: COLORS.present,
    5: COLORS.absent,
    6: COLORS.leave,
    7: COLORS.halfDay,
    8: COLORS.holiday,
    9: COLORS.weekOff,
  };

  for (let row = 1; row < totalRows; row++) {
    Object.entries(summaryColors).forEach(([column, background]) => {
      const cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: Number(column),
      });

      const cell = worksheet[cellAddress];

      if (!cell) return;

      cell.s = {
        font: {
          bold: true,
          sz: 10,
          color: {
            rgb: COLORS.black,
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
    });
  }
};


// DAY STYLE
export const applyDayStyle = (worksheet, totalRows) => {
  for (let row = 1; row < totalRows; row++) {
    for (let col = 10; col < 41; col++) {
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
            rgb: COLORS.black,
          },
        },

        fill: {
          patternType: "solid",

          fgColor: {
            rgb: row % 2 === 0 ? "F9FBFD" : COLORS.white,
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
  }
};


// STATUS STYLE
export const applyStatusStyle = (worksheet, totalRows) => {
  const statusColors = {
    P: {
      background: COLORS.present,
      text: "38761D",
    },

    A: {
      background: COLORS.absent,
      text: "990000",
    },

    L: {
      background: COLORS.leave,
      text: "7F6000",
    },

    HD: {
      background: COLORS.halfDay,
      text: "C65911",
    },

    HO: {
      background: COLORS.holiday,
      text: "1F4E78",
    },

    WO: {
      background: COLORS.weekOff,
      text: "666666",
    },
  };

  for (let row = 1; row < totalRows; row++) {
    for (let col = 10; col < 41; col++) {
      const cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: col,
      });
      const cell = worksheet[cellAddress];
      if (!cell) continue;
      const status = cell.v;
      const style = statusColors[status];
      if (!style) continue;
      cell.s = {
        font: {
          bold: true,
          sz: 10,

          color: {
            rgb: style.text,
          },
        },

        fill: {
          patternType: "solid",

          fgColor: {
            rgb: style.background,
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
  }
};


// COLUMN WIDTH
export const applyColumnWidth = (worksheet, header, rows) => {
  const columnWidths = [];

  for (let col = 0; col < header.length; col++) {
    let maxLength = String(header[col]).length;

    rows.forEach((row) => {
      const value = row[col];

      if (value !== undefined && value !== null) {
        maxLength = Math.max(maxLength, String(value).length);
      }
    });

    let width;

    if (col === 0) {
      width = 12;
    }
    else if (col === 1) {
      width = 22;
    }
    else if (col === 2) {
      width = 20;
    }
    else if (col === 3) {
      width = 20;
    }
    else if (col >= 4 && col <= 9) {
      width = 11;
    }
    else {
      width = 10;
    }

    width = Math.max(width, Math.min(maxLength + 2, 20));
    columnWidths.push({
      wch: width,
    });
  }

  worksheet["!cols"] = columnWidths;
};


// ROW HEIGHT
export const applyRowHeight = (worksheet, totalRows) => {
  worksheet["!rows"] = [];

  // Header
  worksheet["!rows"].push({
    hpt: 32,
  });

  // Employee rows
  for (let row = 1; row < totalRows; row++) {
    worksheet["!rows"].push({
      hpt: 24,
    });
  }
};
