import * as XLSX from "xlsx-js-style";

import {
  getStatus,
  formatDuration,
} from "../AttendencePDF/pdfHelper";

import {
  applyBorderToAllCells,
  applyEmployeeHeaderStyle,
  applyHeaderStyle,
  applySummaryDataStyle,
  applyDayHeaderStyle,
  applyLabelStyle,
  applyStatusStyle,
} from "./ExcelStyle";


const TOTAL_DAYS = 31;
const TOTAL_COLUMNS = TOTAL_DAYS + 1;

const EMPTY_VALUE = "-";


// SAFE EMPLOYEE NAME
const getEmployeeName = (employee) => {
  const firstName =
    employee?.first_name?.trim() || "---";

  const lastName =
    employee?.last_name?.trim() || "---";

  const fullName =
    `${firstName} ${lastName}`.trim();

  return fullName || "---";
};


// GET ATTENDANCE
const getAttendance = (
  attendances,
  index
) => {
  return attendances?.[index] ?? null;
};


// GET DAY VALUES
const createDayValues = (
  attendances,
  callback
) => {
  return Array.from(
    { length: TOTAL_DAYS },
    (_, index) => {
      const attendance =
        getAttendance(
          attendances,
          index
        );

      return callback(
        attendance,
        index
      );
    }
  );
};


// STATUS SUMMARY
const calculateStatusSummary = (
  attendances
) => {
  const summary = {
    P: 0,
    A: 0,
    L: 0,
    HD: 0,
    HO: 0,
    WO: 0,
  };

  attendances.forEach((attendance) => {
    const status =
      getStatus(attendance);

    if (
      Object.prototype.hasOwnProperty.call(
        summary,
        status
      )
    ) {
      summary[status]++;
    }
  });

  return summary;
};


// CREATE EMPLOYEE INFORMATION
const createEmployeeInfoRows = (
  employee
) => {
  const name =
    getEmployeeName(employee);

  return [
    [
      `EmpCode : ${
        employee?.id ?? "---"
      }`,
    ],

    [
      `Name : ${name}`,
    ],

    [
      `Department : ${
        employee?.department_name ||
        "---"
      }`,
    ],

    [
      `Designation : ${
        employee?.designation ||
        "---"
      }`,
    ],

    [],
  ];
};


// CREATE SUMMARY ROWS
const createSummaryRows = (
  employee,
  summary
) => {
  const name =
    getEmployeeName(employee);

  return [
    [
      "EmpCode",
      "Name",
      "Present",
      "Absent",
      "Leave",
      "Halfday",
      "Holiday",
      "WeekOff",
    ],

    [
      employee?.id ?? "---",
      name,
      summary.P,
      summary.A,
      summary.L,
      summary.HD,
      summary.HO,
      summary.WO,
    ],

    [],
  ];
};


// CREATE DAY HEADER
const createDayHeader = () => {
  return [
    "Label",
    ...Array.from(
      { length: TOTAL_DAYS },
      (_, index) => index + 1
    ),
  ];
};


// CREATE ATTENDANCE ROWS
const createAttendanceRows = (
  attendances
) => {
  const inTimeRow = [
    "IN Time",

    ...createDayValues(
      attendances,
      (attendance) =>
        attendance?.in_formatted_time ||
        EMPTY_VALUE
    ),
  ];

  const outTimeRow = [
    "OUT Time",

    ...createDayValues(
      attendances,
      (attendance) =>
        attendance?.out_formatted_time ||
        EMPTY_VALUE
    ),
  ];

  const workingRow = [
    "Working",

    ...createDayValues(
      attendances,
      (attendance) =>
        attendance
          ? formatDuration(
              attendance?.duration || 0
            )
          : EMPTY_VALUE
    ),
  ];

  const overtimeRow = [
    "O.Times",

    ...createDayValues(
      attendances,
      (attendance) =>
        attendance
          ? formatDuration(
              attendance?.ot || 0
            )
          : EMPTY_VALUE
    ),
  ];

  const statusRow = [
    "Status",

    ...createDayValues(
      attendances,
      (attendance) =>
        attendance
          ? getStatus(attendance)
          : EMPTY_VALUE
    ),
  ];

  return [
    inTimeRow,
    outTimeRow,
    workingRow,
    overtimeRow,
    statusRow,
  ];
};


// CREATE ALL ROWS
const createWorksheetRows = (
  employee,
  attendances
) => {
  const summary =
    calculateStatusSummary(
      attendances
    );

  const employeeRows =
    createEmployeeInfoRows(
      employee
    );

  const summaryRows =
    createSummaryRows(
      employee,
      summary
    );

  const dayHeader =
    createDayHeader();

  const attendanceRows =
    createAttendanceRows(
      attendances
    );

  return [
    ...employeeRows,
    ...summaryRows,
    dayHeader,
    ...attendanceRows,
  ];
};


// COLUMN WIDTH
const applyColumnWidths = (
  worksheet,
  rows
) => {
  const widths = [];

  for (
    let col = 0;
    col < TOTAL_COLUMNS;
    col++
  ) {
    let maxLength = 0;

    rows.forEach((row) => {
      const value = row?.[col];

      if (
        value !== undefined &&
        value !== null
      ) {
        maxLength = Math.max(
          maxLength,
          String(value).length
        );
      }
    });

    let width =
      Math.max(
        maxLength + 2,
        col === 0 ? 15 : 10
      );

    width =
      Math.min(width, 22);

    widths.push({
      wch: width,
    });
  }

  worksheet["!cols"] = widths;
};


// MERGE EMPLOYEE INFORMATION
const applyEmployeeMerges = (
  worksheet
) => {
  worksheet["!merges"] =
    Array.from(
      { length: 4 },
      (_, row) => ({
        s: {
          r: row,
          c: 0,
        },

        e: {
          r: row,
          c: TOTAL_COLUMNS - 1,
        },
      })
    );
};


// ROW HEIGHT
const applyRowHeights = (
  worksheet
) => {
  worksheet["!rows"] = [
    { hpt: 24 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 8 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 8 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 22 },
    { hpt: 22 },
  ];
};


// APPLY STYLES
const applyWorksheetStyles = (
  worksheet,
  rows
) => {
  applyBorderToAllCells(
    worksheet
  );

  applyEmployeeHeaderStyle(
    worksheet
  );

  applyHeaderStyle(
    worksheet,
    5,
    8
  );

  applySummaryDataStyle(
    worksheet,
    6,
    8
  );

  // Day header
  applyDayHeaderStyle(
    worksheet,
    8,
    TOTAL_COLUMNS
  );

  // Labels
  applyLabelStyle(
    worksheet,
    [8, 9, 10, 11, 12, 13]
  );

  // Status
  const statusRow =
    rows.length - 1;

  applyStatusStyle(
    worksheet,
    statusRow,
    TOTAL_DAYS
  );

  // Row heights
  applyRowHeights(
    worksheet
  );
};


// SHEET NAME
const createUniqueSheetName = (
  workbook,
  employee,
  employeeIndex
) => {
  const baseName =
    getEmployeeName(employee)
      .replace(
        /[\\/?*[\]:]/g,
        ""
      )
      .substring(0, 25) ||
    `Employee${employeeIndex + 1}`;

  let sheetName =
    baseName;

  let counter = 1;

  while (
    workbook.SheetNames.includes(
      sheetName
    )
  ) {
    const suffix =
      `_${counter}`;

    sheetName =
      `${baseName.substring(
        0,
        31 - suffix.length
      )}${suffix}`;

    counter++;
  }

  return sheetName;
};


// CREATE WORKSHEET
const createEmployeeWorksheet = (
  workbook,
  empData,
  employeeIndex
) => {
  const employee =
    empData?.employee;

  const attendances =
    empData?.attendances ?? [];

  const rows =
    createWorksheetRows(
      employee,
      attendances
    );

  const worksheet =
    XLSX.utils.aoa_to_sheet(
      rows
    );

  applyColumnWidths(
    worksheet,
    rows
  );

  applyEmployeeMerges(
    worksheet
  );

  applyWorksheetStyles(
    worksheet,
    rows
  );

  // Freeze panes
  worksheet["!freeze"] = {
    xSplit: 1,
    ySplit: 9,
  };

  // Sheet name
  const sheetName =
    createUniqueSheetName(
      workbook,
      employee,
      employeeIndex
    );

  // Add worksheet
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    sheetName
  );
};


export const downloadExcel = (
  selectedData = []
) => {
  if (
    !Array.isArray(selectedData) ||
    selectedData.length === 0
  ) {
    console.warn(
      "No employee data available"
    );

    return;
  }

  const workbook =
    XLSX.utils.book_new();

  selectedData.forEach(
    (empData, employeeIndex) => {
      createEmployeeWorksheet(
        workbook,
        empData,
        employeeIndex
      );
    }
  );

  XLSX.writeFile(
    workbook,
    "attendance-report.xlsx"
  );
};