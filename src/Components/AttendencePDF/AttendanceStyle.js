import { StyleSheet } from "@react-pdf/renderer";

export const styles = StyleSheet.create({
  page: {
    padding: 15,
    fontSize: 8,
    backgroundColor: "#ffffff",
    color: "#000000",
  },

  employeeSection: {
    marginBottom: 18,
  },

  /* =========================
     EMPLOYEE HEADER
  ========================= */

  header: {
    borderWidth: 1,
    borderColor: "#000000",
    padding: 6,
    marginBottom: 8,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },

  headerText: {
    fontSize: 9,
    marginBottom: 2,
    color: "#000000",
    fontWeight: "bold",
  },

  employeeName: {
    fontSize: 10,
    fontWeight: "bold",
    color: "#000000",
  },

  designation: {
    fontSize: 9,
    fontWeight: "bold",
    color: "#000000",
  },

  /* =========================
     SUMMARY
  ========================= */

  summaryRow: {
    flexDirection: "row",
    borderWidth: 1,
    borderColor: "#000000",
    marginBottom: 10,
    width: "100%",
  },

  summaryCell: {
    flex: 1,
    borderRightWidth: 1,
    borderColor: "#000000",
    minWidth: 0,
  },

  summaryCellLast: {
    borderRightWidth: 0,
  },

  summaryHeader: {
    backgroundColor: "#333232",
    paddingVertical: 4,
    paddingHorizontal: 3,
    borderBottomWidth: 1,
    borderColor: "#000000",
    alignItems: "center",
    justifyContent: "center",
    minHeight: 22,
  },

  summaryValue: {
    paddingVertical: 4,
    paddingHorizontal: 3,
    alignItems: "center",
    justifyContent: "center",
    minHeight: 22,
  },

  summaryText: {
    fontSize: 9,
    textAlign: "center",
    fontWeight: "bold",
    color: "#FFFFFF",
  },

  summaryValueText: {
    fontSize: 8,
    textAlign: "center",
    color: "#000000",
    fontWeight: "bold",
  },

  /* =========================
     TABLE
  ========================= */

  table: {
    borderWidth: 1,
    borderColor: "#000000",
    width: "100%",
  },

  tableHeaderRow: {
    flexDirection: "row",
    backgroundColor: "#E7E6E6",
    minHeight: 20,
  },

  tableRow: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderColor: "#000000",
    minHeight: 20,
  },

  tableRowLast: {
    borderBottomWidth: 0,
  },

  /* =========================
     LABEL COLUMN
  ========================= */

  cellLabel: {
    width: 70,
    minWidth: 70,
    maxWidth: 70,
    borderRightWidth: 1,
    borderColor: "#000000",
    justifyContent: "center",
    alignItems: "center",
    paddingVertical: 3,
    paddingHorizontal: 2,
  },

  headerLabelCell: {
    backgroundColor: "#E7E6E6",
  },

  /* =========================
     DAY CELLS
  ========================= */

  dayHeaderCell: {
    width: 23,
    minWidth: 23,
    maxWidth: 23,
    borderRightWidth: 1,
    borderColor: "#000000",
    justifyContent: "center",
    alignItems: "center",
    paddingVertical: 3,
    paddingHorizontal: 1,
  },

  cell: {
    width: 23,
    minWidth: 23,
    maxWidth: 23,
    borderRightWidth: 1,
    borderColor: "#000000",
    justifyContent: "center",
    alignItems: "center",
    paddingVertical: 3,
    paddingHorizontal: 1,
  },

  cellLast: {
    borderRightWidth: 0,
  },

  /* =========================
     TEXT
  ========================= */

  tableHeaderText: {
    fontSize: 7,
    fontWeight: "bold",
    color: "#000000",
    textAlign: "center",
  },

  rowLabelText: {
    fontSize: 7,
    fontWeight: "bold",
    color: "#000000",
    textAlign: "center",
  },

  cellText: {
    fontSize: 7,
    textAlign: "center",
    color: "#000000",
    fontWeight: "bold",
    lineHeight: 1.1,
  },

  statusText: {
    fontSize: 6,
    fontWeight: "bold",
    color: "#000000",
    textAlign: "center",
  },

  /* =========================
     STATUS COLORS
  ========================= */

  statusPresent: { backgroundColor: "#C6EFCE" },
  statusAbsent: { backgroundColor: "#FFC7CE" },
  statusLeave: { backgroundColor: "#FFE699" },
  statusHalfDay: { backgroundColor: "#F4CCCC" },
  statusHoliday: { backgroundColor: "#D9EAD3" },
  statusWeekOff: { backgroundColor: "#D9D9D9" },
});
