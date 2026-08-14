import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  getStatus,
  formatDuration,
} from "../AttendencePDF/pdfHelper";
import {
  applyEmployeeInfoStyle,
  applySummaryHeaderStyle,
  applySummaryDataStyle,
  applyDayHeaderStyle,
  applyLabelStyle,
  applyStatusStyle,
  applyAllBorders,
  applyColumnWidths,
  applyRowHeights,
  applyWorksheetSettings,
} from "./ExcelStyle";

const TOTAL_DAYS = 31;
const TOTAL_COLUMNS = TOTAL_DAYS + 1;

// GET EMPLOYEE NAME
const getEmployeeName = (
  employee
) => {
  const firstName =
    employee?.first_name?.trim() || "---";

  const lastName =
    employee?.last_name?.trim() || "---";

  return (
    `${firstName} ${lastName}`.trim() ||
    "---"
  );
};

// STATUS SUMMARY
const getStatusSummary = (
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

  attendances.forEach(
    (attendance) => {
      const status =
        getStatus(attendance);

      if (
        summary[status] !== undefined
      ) {
        summary[status]++;
      }
    }
  );

  return summary;
};

// CREATE EMPLOYEE INFORMATION
const addEmployeeInformation = (
  worksheet,
  employee
) => {
  const name =
    getEmployeeName(employee);

  worksheet.mergeCells(
    "A1:AF1"
  );

  worksheet.mergeCells(
    "A2:AF2"
  );

  worksheet.mergeCells(
    "A3:AF3"
  );

  worksheet.mergeCells(
    "A4:AF4"
  );

  worksheet.getCell(1, 1).value =
    `EmpCode : ${
      employee?.id ?? "---"
    }`;

  worksheet.getCell(2, 1).value =
    `Name : ${name}`;

  worksheet.getCell(3, 1).value =
    `Department : ${
      employee?.department_name ||
      "---"
    }`;

  worksheet.getCell(4, 1).value =
    `Designation : ${
      employee?.designation ||
      "---"
    }`;
};

// ADD SUMMARY
const addSummary = (
  worksheet,
  employee,
  summary
) => {
  const name = getEmployeeName(employee);

  worksheet.addRow([]);

  worksheet.addRow([
    "EmpCode",
    "Name",
    "Present",
    "Absent",
    "Leave",
    "Halfday",
    "Holiday",
    "WeekOff",
  ]);

  // Data
  worksheet.addRow([
    employee?.id ?? "---",
    name,
    summary.P,
    summary.A,
    summary.L,
    summary.HD,
    summary.HO,
    summary.WO,
  ]);
};

// ADD ATTENDANCE DATA
const addAttendanceData = (
  worksheet,
  attendances
) => {
  worksheet.addRow([]);

  worksheet.addRow([
    "Label",

    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) =>
        index + 1
    ),
  ]);

  // IN TIME
  worksheet.addRow([
    "IN Time",
    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) =>
        attendances[index]
          ?.in_formatted_time ||
        "-",
    ),
  ]);

  // OUT TIME
  worksheet.addRow([
    "OUT Time",

    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) =>
        attendances[index]
          ?.out_formatted_time ||
        "-",
    ),
  ]);

  // WORKING
  worksheet.addRow([
    "Working",
    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) => {
        const attendance =
          attendances[index];

        return attendance
          ? formatDuration(
              attendance.duration ||
                0
            )
          : "-";
      }
    ),
  ]);

  // OVERTIME
  worksheet.addRow([
    "O.Times",

    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) => {
        const attendance =
          attendances[index];

        return attendance
          ? formatDuration(
              attendance.ot || 0
            )
          : "-";
      }
    ),
  ]);

  // STATUS
  worksheet.addRow([
    "Status",

    ...Array.from(
      {
        length: TOTAL_DAYS,
      },
      (_, index) => {
        const attendance =
          attendances[index];

        return attendance
          ? getStatus(attendance)
          : "-";
      }
    ),
  ]);
};

// APPLY ALL STYLES
const applyStyles = (
  worksheet
) => {
  applyEmployeeInfoStyle(
    worksheet,
    TOTAL_COLUMNS
  );
  applySummaryHeaderStyle(
    worksheet,
    8
  );
  applySummaryDataStyle(
    worksheet,
    8
  );

  // Day header
  applyDayHeaderStyle(
    worksheet,
    TOTAL_COLUMNS
  );

  // Label column
  applyLabelStyle(
    worksheet,
    9,
    14
  );

  // Status colors
  applyStatusStyle(
    worksheet,
    14,
    TOTAL_DAYS
  );

  // General borders
  applyAllBorders(
    worksheet
  );

  // Column widths
  applyColumnWidths(
    worksheet,
    TOTAL_COLUMNS
  );

  // Row heights
  applyRowHeights(
    worksheet
  );

  // Freeze + page setup
  applyWorksheetSettings(
    worksheet
  );
};

// CREATE UNIQUE SHEET NAME
const getUniqueSheetName = (
  workbook,
  employee,
  index
) => {
  const name =
    getEmployeeName(employee);

  let sheetName =
    name
      .replace(
        /[\\/?*[\]:]/g,
        ""
      )
      .substring(0, 25) ||
    `Employee${index + 1}`;

  let counter = 1;

  while (
    workbook.getWorksheet(
      sheetName
    )
  ) {
    const suffix =
      `_${counter}`;

    sheetName =
      `${name.substring(
        0,
        31 - suffix.length
      )}${suffix}`;

    counter++;
  }

  return sheetName;
};

// CREATE EMPLOYEE SHEET
const createEmployeeSheet = (
  workbook,
  empData,
  index
) => {
  const employee =
    empData?.employee;

  const attendances =
    empData?.attendances ?? [];

  const summary =
    getStatusSummary(
      attendances
    );

  const sheetName =
    getUniqueSheetName(
      workbook,
      employee,
      index
    );

  const worksheet =
    workbook.addWorksheet(
      sheetName
    );

  addEmployeeInformation(
    worksheet,
    employee
  );

  addSummary(
    worksheet,
    employee,
    summary
  );

  addAttendanceData(
    worksheet,
    attendances
  );

  applyStyles(
    worksheet
  );
};

// DOWNLOAD EXCEL

export const downloadExcel = async (
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

  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Attendance System";
    workbook.lastModifiedBy = "Attendance System";
    workbook.created = new Date();
    workbook.modified = new Date();

    // Create employee sheets
    selectedData.forEach(
      (empData, index) => {
        createEmployeeSheet(
          workbook,
          empData,
          index
        );
      }
    );

    // Generate buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Create Blob
    const blob = new Blob(
      [buffer],
      {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    );

    // Download
    saveAs(
      blob,
      "attendance-report.xlsx"
    );
  } catch (error) {
    console.error(
      "Excel generation failed:",
      error
    );
  }
};