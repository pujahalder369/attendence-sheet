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
} from "./OnlyExcelStyle";

export const downloadOnlyExcel = (selectedData) => {
  const workbook = XLSX.utils.book_new();

  selectedData.forEach((empData, employeeIndex) => {
    const employee = empData?.employee;
    const attendances = empData?.attendances ?? [];

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
        employee?.department_name ?? "---"
      }`,
    ]);

    rows.push([
      `Designation : ${
        employee?.designation ?? "---"
      }`,
    ]);

    // Empty row
    rows.push([]);

  
    // STATUS CALCULATION
  

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

  
    // SUMMARY HEADER
  

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

  
    // SUMMARY DATA
  

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

    // Empty row
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
            ? formatDuration(att?.duration || 0)
            : "-";
        }
      ),
    ]);

  
    // OVERTIME
  

    rows.push([
      "O.Times",

      ...Array.from(
        { length: 31 },
        (_, index) => {
          const att = attendances[index];

          return att
            ? formatDuration(att?.ot || 0)
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
  

    const columnWidths = [];

    for (let col = 0; col < 32; col++) {
      let maxLength = 0;

      rows.forEach((row) => {
        const value = row[col];

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

      let width = Math.max(
        maxLength + 2,
        10
      );

      // Label column
      if (col === 0) {
        width = Math.max(
          maxLength + 3,
          15
        );
      }

      // Day columns
      if (col >= 1) {
        width = Math.max(
          width,
          11
        );
      }

      // Maximum width
      width = Math.min(
        width,
        22
      );

      columnWidths.push({
        wch: width,
      });
    }

    worksheet["!cols"] =
      columnWidths;

  
    // MERGE EMPLOYEE INFORMATION
  

    worksheet["!merges"] = [
      {
        s: {
          r: 0,
          c: 0,
        },
        e: {
          r: 0,
          c: 31,
        },
      },

      {
        s: {
          r: 1,
          c: 0,
        },
        e: {
          r: 1,
          c: 31,
        },
      },

      {
        s: {
          r: 2,
          c: 0,
        },
        e: {
          r: 2,
          c: 31,
        },
      },

      {
        s: {
          r: 3,
          c: 0,
        },
        e: {
          r: 3,
          c: 31,
        },
      },
    ];

  
    // GENERAL BORDER
  

    applyBorderToAllCells(
      worksheet
    );

  
    // EMPLOYEE HEADER
  

    applyEmployeeHeaderStyle(
      worksheet
    );

  
    // SUMMARY HEADER
    // ROW 5
  

    applyHeaderStyle(
      worksheet,
      5,
      8
    );

  
    // SUMMARY DATA
    // ROW 6
  

    applySummaryDataStyle(
      worksheet,
      6,
      8
    );

  
    // DAY HEADER
    // ROW 8
  

    applyDayHeaderStyle(
      worksheet,
      7,
      32
    );

  
    // LABEL COLUMN
  

    applyLabelStyle(
      worksheet,
      [
        7,
        8,
        9,
        10,
        11,
        12,
      ]
    );

  
    // STATUS COLORS
  

    const statusRow =
      rows.length - 1;

    applyStatusStyle(
      worksheet,
      statusRow
    );

  
    // ROW HEIGHT
  

    worksheet["!rows"] = [
      { hpt: 22 }, // EmpCode
      { hpt: 22 }, // Name
      { hpt: 22 }, // Department
      { hpt: 22 }, // Designation
      { hpt: 8 },  // Empty
      { hpt: 22 }, // Summary header
      { hpt: 22 }, // Summary data
      { hpt: 8 },  // Empty
      { hpt: 22 }, // Day header
      { hpt: 22 }, // In
      { hpt: 22 }, // Out
      { hpt: 22 }, // Working
      { hpt: 22 }, // OT
      { hpt: 22 }, // Status
    ];

  
    // FREEZE
  

    worksheet["!freeze"] = {
      xSplit: 1,
      ySplit: 8,
    };

  
    // SHEET NAME
  

    let sheetName =
      `${employee?.first_name ?? "Employee"}`
        .replace(
          /[\\/?*[\]:]/g,
          ""
        )
        .substring(0, 25);

    if (!sheetName) {
      sheetName =
        `Employee${employeeIndex + 1}`;
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

  
    // ADD SHEET
  

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      sheetName
    );
  });


  // DOWNLOAD


  XLSX.writeFile(
    workbook,
    "attendance-report.xlsx"
  );
};