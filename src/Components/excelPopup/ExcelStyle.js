import * as XLSX from "xlsx-js-style";

// BORDER

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


// CENTER


export const centerStyle = {
  horizontal: "center",
  vertical: "center",
  wrapText: true,
};


// DARK HEADER


export const headerStyle = {
  font: {
    bold: true,
    color: { rgb: "FFFFFF" },
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


// EMPLOYEE HEADER


export const employeeHeaderStyle = {
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
    wrapText: false,
  },

  border: borderStyle,
};


// STATUS COLORS


export const statusColors = {
  P: "4CAF50", // Green
  A: "F44336", // Red
  L: "FFC107", // Yellow
  HD: "FF7043", // Orange
  HO: "42A5F5", // Blue
  WO: "9E9E9E", // Grey
};


// APPLY BORDER


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


// EMPLOYEE HEADER STYLE


export const applyEmployeeHeaderStyle = (worksheet) => {
  ["A1", "A2", "A3", "A4"].forEach((cellAddress) => {
    if (worksheet[cellAddress]) {
      worksheet[cellAddress].s = {
        ...employeeHeaderStyle,
      };
    }
  });
};


// HEADER STYLE


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

    if (worksheet[cellAddress]) {
      worksheet[cellAddress].s = {
        ...headerStyle,
      };
    }
  }
};


// STATUS STYLE


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