import ExcelJS from "exceljs";

export const COLORS = {
  employeeHeader: "FFD9EAF7",
  summaryHeader: "FF404040",
  summaryData: "FFF3F6F9",
  dayHeader: "FF5B9BD5",
  label: "FFEAF2F8",
  border: "FFB7B7B7",
  white: "FFFFFFFF",
  black: "FF000000",
};

export const STATUS_COLORS = {
  P: {
    background: "FFC6EFCE",
    font: "FF006100",
  },

  A: {
    background: "FFFFC7CE",
    font: "FF9C0006",
  },

  L: {
    background: "FFFFEB9C",
    font: "FF9C6500",
  },

  HD: {
    background: "FFF4B084",
    font: "FF843C0C",
  },

  HO: {
    background: "FF9DC3E6",
    font: "FF1F4E78",
  },

  WO: {
    background: "FFD9E1F2",
    font: "FF404040",
  },
};

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

export const centerAlignment = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

export const leftAlignment = {
  horizontal: "left",
  vertical: "middle",
  wrapText: true,
};

export const applyEmployeeInfoStyle = (
  worksheet,
  totalColumns
) => {
  for (let row = 1; row <= 4; row++) {
    const cell =
      worksheet.getCell(row, 1);

    cell.font = {
      bold: true,
      size: 11,
      color: {
        argb: "FF222222",
      },
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: COLORS.employeeHeader,
      },
    };

    cell.alignment = {
      ...leftAlignment,
    };

    applyBorderToRange(
      worksheet,
      row,
      row,
      1,
      totalColumns
    );
  }
};

export const applySummaryHeaderStyle = (
  worksheet,
  totalColumns = 8
) => {
  for (
    let col = 1;
    col <= totalColumns;
    col++
  ) {
    const cell =
      worksheet.getCell(6, col);

    cell.font = {
      bold: true,
      size: 11,
      color: {
        argb: COLORS.white,
      },
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: COLORS.summaryHeader,
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

export const applySummaryDataStyle = (
  worksheet,
  totalColumns = 8
) => {
  for (
    let col = 1;
    col <= totalColumns;
    col++
  ) {
    const cell =
      worksheet.getCell(7, col);

    cell.font = {
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
        argb: COLORS.summaryData,
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

export const applyDayHeaderStyle = (
  worksheet,
  totalColumns
) => {
  for (
    let col = 1;
    col <= totalColumns;
    col++
  ) {
    const cell =
      worksheet.getCell(9, col);

    cell.font = {
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
        argb: COLORS.dayHeader,
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

export const applyLabelStyle = (
  worksheet,
  startRow,
  endRow
) => {
  for (
    let row = startRow;
    row <= endRow;
    row++
  ) {
    const cell =
      worksheet.getCell(row, 1);

    cell.font = {
      bold: true,
      size: 10,
      color: {
        argb: "FF222222",
      },
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: COLORS.label,
      },
    };

    cell.alignment = {
      ...leftAlignment,
    };

    cell.border = {
      ...borderStyle,
    };
  }
};

export const applyStatusStyle = (
  worksheet,
  statusRow,
  totalDays
) => {
  for (
    let col = 2;
    col <= totalDays + 1;
    col++
  ) {
    const cell =
      worksheet.getCell(
        statusRow,
        col
      );

    const status = cell.value;
    const config = STATUS_COLORS[status];

    if (!config) {
      cell.font = {
        bold: true,
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

      continue;
    }

    cell.font = {
      bold: true,
      size: 10,
      color: {
        argb: config.font,
      },
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: config.background,
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

export const applyAllBorders = (
  worksheet
) => {
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.border = {
        ...borderStyle,
        ...cell.border,
      };
    });
  });
};

export const applyColumnWidths = (
  worksheet,
  totalColumns
) => {
  worksheet.getColumn(1).width = 25;

  for (
    let col = 2;
    col <= totalColumns;
    col++
  ) {
    worksheet.getColumn(col).width = 12;
  }
};

export const applyRowHeights = (
  worksheet
) => {
  worksheet.getRow(1).height = 24;

  for (
    let row = 2;
    row <= 4;
    row++
  ) {
    worksheet.getRow(row).height = 22;
  }

  worksheet.getRow(5).height = 8;
  worksheet.getRow(6).height = 22;
  worksheet.getRow(7).height = 22;
  worksheet.getRow(8).height = 8;

  for (
    let row = 9;
    row <= 14;
    row++
  ) {
    worksheet.getRow(row).height = 22;
  }
};

export const applyWorksheetSettings = (
  worksheet
) => {
  worksheet.views = [
    {
      state: "frozen",
      xSplit: 1,
      ySplit: 9,
    },
  ];

  worksheet.pageSetup = {
    orientation: "landscape",
    paperSize:
      worksheet.PAPERSIZE_A4,
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
};

export const applyBorderToRange = (
  worksheet,
  startRow,
  endRow,
  startCol,
  endCol
) => {
  for (
    let row = startRow;
    row <= endRow;
    row++
  ) {
    for (
      let col = startCol;
      col <= endCol;
      col++
    ) {
      worksheet.getCell(
        row,
        col
      ).border = {
        ...borderStyle,
      };
    }
  }
};