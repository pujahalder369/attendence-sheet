import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
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
  applyWorksheetSettings,
} from "./OnlyExcelStyle";

const TOTAL_DAYS = 31;
const EMPLOYEE_COLUMNS = 9;
const TOTAL_COLUMNS = EMPLOYEE_COLUMNS + TOTAL_DAYS;

export const downloadOnlyExcel = async (selectedData = []) => {
  try {
    if (!Array.isArray(selectedData) || selectedData.length === 0) {
      console.warn("No employee data available");
      return;
    }
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Attendance System";
    workbook.lastModifiedBy = "Attendance System";
    workbook.created = new Date();
    workbook.modified = new Date();
    const worksheet = workbook.addWorksheet("Attendance Report");

    const header = [
      "EmpCode",
      "Name",
      "Department",
      "Present",
      "Absent",
      "Leave",
      "HD",
      "HO",
      "WO",
      ...Array.from({ length: TOTAL_DAYS }, (_, index) => index + 1),
    ];

    worksheet.addRow(header);
    selectedData.forEach((empData) => {
      const employee = empData?.employee;

      const attendances = empData?.attendances ?? [];

      const statusCount = {
        P: 0,
        A: 0,
        L: 0,
        HD: 0,
        HO: 0,
        WO: 0,
      };
      const dayStatuses = Array(TOTAL_DAYS).fill("-");

      attendances.forEach((attendance, index) => {
        const status = getStatus(attendance);
        if (statusCount[status] !== undefined) {
          statusCount[status]++;
        }
        if (index < TOTAL_DAYS) {
          dayStatuses[index] = status;
        }
      });

      const employeeName = `${employee?.first_name ?? "---"} ${
        employee?.last_name ?? "---"
      }`;

      const row = [
        employee?.id ?? "---",
        employeeName,
        employee?.department_name ?? "---",
        statusCount.P,
        statusCount.A,
        statusCount.L,
        statusCount.HD,
        statusCount.HO,
        statusCount.WO,
        ...dayStatuses,
      ];

      worksheet.addRow(row);
    });
    const totalRows = worksheet.rowCount;

    applyGeneralStyle(worksheet, totalRows, TOTAL_COLUMNS);
    applyHeaderStyle(worksheet, TOTAL_COLUMNS);
    applyEmployeeStyle(worksheet, totalRows);
    applySummaryStyle(worksheet, totalRows);
    applyDayStyle(worksheet, totalRows);
    applyStatusStyle(worksheet, totalRows);
    applyColumnWidth(worksheet, header);
    applyRowHeight(worksheet, totalRows);
    applyWorksheetSettings(worksheet);

    worksheet.views = [
      {
        state: "frozen",
        xSplit: 2,
        ySplit: 1,
      },
    ];

    worksheet.autoFilter = {
      from: {
        row: 1,
        column: 1,
      },

      to: {
        row: totalRows,
        column: TOTAL_COLUMNS,
      },
    };

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "attendance-report.xlsx");
  } catch (error) {
    console.error("Excel generation failed:", error);
  }
};
