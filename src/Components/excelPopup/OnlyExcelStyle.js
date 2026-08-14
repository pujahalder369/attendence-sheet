export const COLORS = {
  header: "FF17365D",
  dayHeader: "FF2F75B5",
  white: "FFFFFFFF",
  black: "FF000000",
  border: "FFD9E1F2",
  employeeEven: "FFF8FAFC",
  employeeOdd: "FFEEF4FB",
  present: "FFC6EFCE",
  absent: "FFFFC7CE",
  leave: "FFFFEB9C",
  halfDay: "FFF4B183",
  holiday: "FF9DC3E6",
  weekOff: "FFD9D9D9",
};

// BORDER
export const borderStyle = {
  top: {
    style: "thin",
    color: {
      argb: COLORS.border,
    },
  },

  bottom: {
    style: "thin",
    color: {
      argb: COLORS.border,
    },
  },

  left: {
    style: "thin",
    color: {
      argb: COLORS.border,
    },
  },

  right: {
    style: "thin",
    color: {
      argb: COLORS.border,
    },
  },
};

// CENTER ALIGNMENT
export const centerAlignment = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

// GENERAL STYLE
export const applyGeneralStyle = (worksheet, totalRows, totalColumns) => {
  for (let row = 1; row <= totalRows; row++) {
    for (let col = 1; col <= totalColumns; col++) {
      const cell = worksheet.getCell(row, col);

      cell.font = {
        name: "Calibri",
        size: 10,
        color: {
          argb: COLORS.black,
        },
      };

      cell.alignment = {
        ...centerAlignment,
      };

      cell.border = {
        ...borderStyle,
      };
    }
  }
};

// HEADER STYLE
export const applyHeaderStyle = (worksheet, totalColumns) => {
  for (let col = 1; col <= totalColumns; col++) {
    const cell = worksheet.getCell(1, col);

    cell.font = {
      name: "Calibri",
      bold: true,
      size: 10,
      color: {
        argb: COLORS.white,
      },
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: col <= 10 ? COLORS.header : COLORS.dayHeader,
      },
    };

    cell.alignment = {
      ...centerAlignment,
    };

    cell.border = {
      ...borderStyle,
    };
  }
};

// EMPLOYEE STYLE
export const applyEmployeeStyle = (worksheet, totalRows) => {
  for (let row = 2; row <= totalRows; row++) {
    for (let col = 1; col <= 4; col++) {
      const cell = worksheet.getCell(row, col);

      cell.font = {
        name: "Calibri",
        bold: true,
        size: 10,
        color: {
          argb: COLORS.black,
        },
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: row % 2 === 0 ? COLORS.employeeEven : COLORS.employeeOdd,
        },
      };

      cell.alignment = {
        horizontal: col === 1 ? "center" : "left",

        vertical: "middle",

        wrapText: true,
      };

      cell.border = {
        ...borderStyle,
      };
    }
  }
};

// SUMMARY STYLE
export const applySummaryStyle = (worksheet, totalRows) => {
  const summaryColors = {
    5: COLORS.present,
    6: COLORS.absent,
    7: COLORS.leave,
    8: COLORS.halfDay,
    9: COLORS.holiday,
    10: COLORS.weekOff,
  };

  for (let row = 2; row <= totalRows; row++) {
    Object.entries(summaryColors).forEach(([column, background]) => {
      const cell = worksheet.getCell(row, Number(column));

      cell.font = {
        name: "Calibri",
        bold: true,
        size: 10,
        color: {
          argb: COLORS.black,
        },
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: background,
        },
      };

      cell.alignment = {
        ...centerAlignment,
      };

      cell.border = {
        ...borderStyle,
      };
    });
  }
};

// DAY STYLE
export const applyDayStyle = (worksheet, totalRows) => {
  for (let row = 2; row <= totalRows; row++) {
    for (let col = 11; col <= 41; col++) {
      const cell = worksheet.getCell(row, col);

      cell.font = {
        name: "Calibri",
        bold: true,
        size: 10,
        color: {
          argb: COLORS.black,
        },
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: row % 2 === 0 ? "FFF9FBFD" : COLORS.white,
        },
      };

      cell.alignment = {
        ...centerAlignment,
      };

      cell.border = {
        ...borderStyle,
      };
    }
  }
};

// STATUS STYLE
export const applyStatusStyle = (worksheet, totalRows) => {
  const statusColors = {
    P: {
      background: COLORS.present,
      text: "FF38761D",
    },

    A: {
      background: COLORS.absent,
      text: "FF990000",
    },

    L: {
      background: COLORS.leave,
      text: "FF7F6000",
    },

    HD: {
      background: COLORS.halfDay,
      text: "FFC65911",
    },

    HO: {
      background: COLORS.holiday,
      text: "FF1F4E78",
    },

    WO: {
      background: COLORS.weekOff,
      text: "FF666666",
    },
  };

  for (let row = 2; row <= totalRows; row++) {
    for (let col = 11; col <= 41; col++) {
      const cell = worksheet.getCell(row, col);
      const status = cell.value;
      const style = statusColors[status];
      if (!style) continue;
      cell.font = {
        name: "Calibri",
        bold: true,
        size: 10,
        color: {
          argb: style.text,
        },
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: style.background,
        },
      };

      cell.alignment = {
        ...centerAlignment,
      };

      cell.border = {
        ...borderStyle,
      };
    }
  }
};

// COLUMN WIDTH
export const applyColumnWidth = (worksheet, header) => {
  for (let col = 1; col <= header.length; col++) {
    let width;

    if (col === 1) {
      width = 14;
    }
    else if (col === 2) {
      width = 24;
    }
    else if (col === 3) {
      width = 22;
    }
    else if (col === 4) {
      width = 22;
    }
    else if (col >= 5 && col <= 10) {
      width = 11;
    }
    else {
      width = 10;
    }

    worksheet.getColumn(col).width = width;
  }
};

// ROW HEIGHT
export const applyRowHeight = (worksheet, totalRows) => {
  worksheet.getRow(1).height = 32;
  for (let row = 2; row <= totalRows; row++) {
    worksheet.getRow(row).height = 28;
  }
};

// WORKSHEET SETTINGS
export const applyWorksheetSettings = (worksheet) => {
  worksheet.pageSetup = {
    orientation: "landscape",
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  worksheet.pageMargins = {
    left: 0.25,
    right: 0.25,
    top: 0.5,
    bottom: 0.5,
    header: 0.2,
    footer: 0.2,
  };

  worksheet.pageSetup.printTitlesRow = "1:1";
  worksheet.pageSetup.horizontalCentered = true;
};
