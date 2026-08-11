import * as XLSX from "xlsx-js-style";
import {
  getStatus,
  formatDuration,
} from "../AttendencePDF/pdfHelper";
import {
  applyBorderToAllCells,
  applyEmployeeHeaderStyle,
  applyHeaderStyle,
  applyStatusStyle,
} from "./ExcelStyle";

export const downloadExcel = (selectedData) => {
  const workbook = XLSX.utils.book_new();

  selectedData.forEach(
    (empData, employeeIndex) => {
      const employee = empData?.employee;

      const attendances =
        empData?.attendances ?? [];

      const rows = [];

      // EMPLOYEE INFORMATION

      rows.push([
        `EmpCode : ${employee?.id ?? "---"}`,
      ]);

      rows.push([
        `Name : ${employee?.first_name ?? "---"} ${
          employee?.last_name ?? "---"
        }`,
      ]);

      rows.push([
        `Department : ${
          employee?.department_name || "---"
        }`,
      ]);

      rows.push([
        `Designation : ${
          employee?.designation || "---"
        }`,
      ]);

      rows.push([]);

      const statusList = attendances.map((att) =>
        getStatus(att)
      );

      const present = statusList.filter(
        (status) => status === "P"
      ).length;

      const absent = statusList.filter(
        (status) => status === "A"
      ).length;

      const leave = statusList.filter(
        (status) => status === "L"
      ).length;

      const halfDay = statusList.filter(
        (status) => status === "HD"
      ).length;

      const holiday = statusList.filter(
        (status) => status === "HO"
      ).length;

      const weekOff = statusList.filter(
        (status) => status === "WO"
      ).length;

      rows.push([
        "EmpCode",
        "Name",
        "Present",
        "Absent",
        "Leave",
        "Halfday",
        "Holiday",
        "WeekOff",
      ]);

      rows.push([
        employee?.id ?? "---",

        `${employee?.first_name ?? "---"} ${
          employee?.last_name ?? "---"
        }`,

        present,
        absent,
        leave,
        halfDay,
        holiday,
        weekOff,
      ]);

      rows.push([]);

      // DAY HEADER

      rows.push([
        "Label",

        ...Array.from(
          { length: 31 },
          (_, index) => index + 1
        ),
      ]);

      // IN TIME

      rows.push([
        "IN Time",

        ...Array.from(
          { length: 31 },
          (_, index) => {
            const att = attendances[index];

            return (
              att?.in_formatted_time || "-"
            );
          }
        ),
      ]);

      // OUT TIME

      rows.push([
        "OUT Time",

        ...Array.from(
          { length: 31 },
          (_, index) => {
            const att = attendances[index];

            return (
              att?.out_formatted_time || "-"
            );
          }
        ),
      ]);

      // WORKING

      rows.push([
        "Working",

        ...Array.from(
          { length: 31 },
          (_, index) => {
            const att = attendances[index];

            return att
              ? formatDuration(
                  att?.duration || 0
                )
              : "-";
          }
        ),
      ]);

      // O.Times

      rows.push([
        "O.Times",

        ...Array.from(
          { length: 31 },
          (_, index) => {
            const att = attendances[index];

            return att
              ? formatDuration(
                  att?.ot || 0
                )
              : "-";
          }
        ),
      ]);

      // STATUS

      rows.push([
        "Status",

        ...Array.from(
          { length: 31 },
          (_, index) => {
            const att = attendances[index];

            return att
              ? getStatus(att)
              : "-";
          }
        ),
      ]);

      // CREATE WORKSHEET

      const worksheet =
        XLSX.utils.aoa_to_sheet(rows);

      // COLUMN WIDTH

      worksheet["!cols"] = [
        { wch: 18 },

        ...Array.from(
          { length: 31 },
          () => ({
            wch: 12,
          })
        ),
      ];

      // MERGE EMPLOYEE HEADER

      worksheet["!merges"] = [
        {
          s: { r: 0, c: 0 },
          e: { r: 0, c: 31 },
        },
        {
          s: { r: 1, c: 0 },
          e: { r: 1, c: 31 },
        },
        {
          s: { r: 2, c: 0 },
          e: { r: 2, c: 31 },
        },
        {
          s: { r: 3, c: 0 },
          e: { r: 3, c: 31 },
        },
      ];

      // GENERAL STYLE

      applyBorderToAllCells(worksheet);

      // EMPLOYEE HEADER

      applyEmployeeHeaderStyle(
        worksheet
      );

      // SUMMARY HEADER

      applyHeaderStyle(
        worksheet,
        5,
        8
      );

      // DAY HEADER

      applyHeaderStyle(
        worksheet,
        7,
        32
      );

      // STATUS COLORS

      const statusRow =
        rows.length - 1;

      applyStatusStyle(
        worksheet,
        statusRow
      );

      // SHEET NAME

      let sheetName =
        `${employee?.first_name ?? "Employee"}`
          .replace(/[\\/?*[\]:]/g, "")
          .substring(0, 25);

      if (!sheetName) {
        sheetName = `Employee${
          employeeIndex + 1
        }`;
      }

      if (
        workbook.SheetNames.includes(
          sheetName
        )
      ) {
        sheetName =
          `${sheetName}_${employeeIndex + 1}`
            .substring(0, 31);
      }

      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        sheetName
      );
    }
  );

  // DOWNLOAD

  XLSX.writeFile(
    workbook,
    "attendance-report.xlsx"
  );
};