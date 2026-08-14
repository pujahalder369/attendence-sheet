import * as XLSX from "xlsx-js-style";
import { getStatus } from "../AttendencePDF/pdfHelper";
import {
  applyGeneralStyle,
  applyHeaderStyle,
  applyEmployeeStyle,
  applySummaryStyle,
  applyDayStyle,
  applyStatusStyle,
  applyColumnWidth,
  applyRowHeight,
} from "./OnlyExcelStyle";

export const downloadOnlyExcel = (selectedData) => {
  const workbook = XLSX.utils.book_new();

  // HEADER
  const header = [
    "EmpCode",
    "Name",
    "Department",
    "Designation",
    "Present",
    "Absent",
    "Leave",
    "HD",
    "HO",
    "WO",
  ];

  // DAY 1 - DAY 31
  for (let day = 1; day <= 31; day++) {
    header.push(day);
  }

  // DATA ROWS
  const rows = [];

  selectedData.forEach((empData) => {
    const employee = empData?.employee;
    const attendances = empData?.attendances ?? [];
    const statusList = attendances.map((attendance) => getStatus(attendance));

    // STATUS COUNT
    const present = statusList.filter((status) => status === "P").length;
    const absent = statusList.filter((status) => status === "A").length;
    const leave = statusList.filter((status) => status === "L").length;
    const halfDay = statusList.filter((status) => status === "HD").length;
    const holiday = statusList.filter((status) => status === "HO").length;
    const weekOff = statusList.filter((status) => status === "WO").length;

    // EMPLOYEE INFORMATION
    const row = [
      employee?.id ?? "---",

      `${employee?.first_name ?? "---"} ${employee?.last_name ?? "---"}`,
      employee?.department_name ?? "---",
      employee?.designation ?? "---",
      present,
      absent,
      leave,
      halfDay,
      holiday,
      weekOff,
    ];

    // DAY 1 - DAY 31 STATUS
    for (let day = 0; day < 31; day++) {
      const attendance = attendances[day];
      const status = attendance ? getStatus(attendance) : "-";
      row.push(status);
    }

    rows.push(row);
  });

  // CREATE WORKSHEET
  const worksheet = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const totalRows = rows.length + 1;
  const totalColumns = header.length;

  // GENERAL STYLE
  applyGeneralStyle(worksheet, totalRows, totalColumns);

  // HEADER STYLE
  applyHeaderStyle(worksheet, totalColumns);

  // EMPLOYEE STYLE
  applyEmployeeStyle(worksheet, totalRows);

  // SUMMARY STYLE
  applySummaryStyle(worksheet, totalRows);

  // DAY STYLE
  applyDayStyle(worksheet, totalRows);

  // STATUS COLOR
  applyStatusStyle(worksheet, totalRows);

  // COLUMN WIDTH
  applyColumnWidth(worksheet, header, rows);

  // ROW HEIGHT
  applyRowHeight(worksheet, totalRows);

  // FREEZE
  worksheet["!freeze"] = {
    xSplit: 10,
    ySplit: 1,
  };

  // AUTO FILTER
  worksheet["!autofilter"] = {
    ref: `A1:${XLSX.utils.encode_col(totalColumns - 1)}${totalRows}`,
  };
  XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance Report");
  XLSX.writeFile(workbook, "attendance-report.xlsx");
};
