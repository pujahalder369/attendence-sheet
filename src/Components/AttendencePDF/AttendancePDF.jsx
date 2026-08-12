import { Document, Page, Text, View } from "@react-pdf/renderer";
import { styles } from "./AttendanceStyle";
import {
    calculateSummary,
    getStatus,
    formatDuration,
} from "./pdfHelper";

const days = Array.from({ length: 31 }, (_, i) => i + 1);

const reportRows = [
    {
        label: "IN Time",
        key: "in_formatted_time",
    },
    {
        label: "OUT Time",
        key: "out_formatted_time",
    },
    {
        label: "Working",
        key: "duration",
        isDuration: true,
    },
    {
        label: "O.Times",
        key: "ot",
        isDuration: true,
    },
    {
        label: "Status",
        key: "status",
    },
];

const AttendancePDF = ({
    data = [],
    selectedIds = [],
}) => {
    const filteredData =
        selectedIds.length > 0
            ? data.filter((emp) =>
                selectedIds.includes(emp?.employee?.id)
            )
            : data;

    return (
        <Document>
            <Page
                size="A4"
                orientation="landscape"
                style={styles.page}
                wrap
            >
                {filteredData.map((emp, index) => {
                    const summaryItems = calculateSummary(
                        emp?.attendances ?? []
                    );

                    return (
                        <View
                            key={index}
                            style={styles.employeeSection}
                            wrap={false}
                        >

                            {/* HEADER */}

                            <View style={styles.header}>
                                <Text style={styles.headerText}>
                                    EmpCode : {emp?.employee?.id ?? "---"}
                                </Text>

                                <Text style={styles.employeeName}>
                                    Name : {emp?.employee?.first_name ?? "---"}{" "}
                                    {emp?.employee?.last_name ?? "---"}
                                </Text>

                                <Text style={styles.headerText}>
                                    Department :{" "}
                                    {emp?.employee?.department_name || "---"}
                                </Text>

                                <Text style={styles.designation}>
                                    Designation :{" "}
                                    {emp?.employee?.designation || "---"}
                                </Text>
                            </View>

                            {/* SUMMARY */}

                            <View style={styles.summaryRow}>
                                {summaryItems.map(
                                    (item, summaryIndex) => (
                                        <View
                                            key={summaryIndex}
                                            style={[
                                                styles.summaryCell,
                                                summaryIndex ===
                                                summaryItems.length - 1 &&
                                                styles.summaryCellLast,
                                            ]}
                                        >
                                            <View
                                                style={
                                                    styles.summaryHeader
                                                }
                                            >
                                                <Text
                                                    style={
                                                        styles.summaryText
                                                    }
                                                >
                                                    {item.label}
                                                </Text>
                                            </View>

                                            <View
                                                style={
                                                    styles.summaryValue
                                                }
                                            >
                                                <Text
                                                    style={
                                                        styles.summaryValueText
                                                    }
                                                >
                                                    {item.value}
                                                </Text>
                                            </View>
                                        </View>
                                    )
                                )}
                            </View>

                            {/* TABLE HEADER */}

                            <View style={styles.table}>
                                <View
                                    style={[
                                        styles.tableRow,
                                        styles.tableHeaderRow,
                                    ]}
                                >
                                    <View
                                        style={[
                                            styles.cellLabel,
                                            styles.headerLabelCell,
                                        ]}
                                    >
                                        <Text
                                            style={
                                                styles.tableHeaderText
                                            }
                                        >
                                            Label
                                        </Text>
                                    </View>

                                    {days.map((day) => (
                                        <View
                                            key={day}
                                            style={
                                                styles.dayHeaderCell
                                            }
                                        >
                                            <Text
                                                style={
                                                    styles.tableHeaderText
                                                }
                                            >
                                                {day}
                                            </Text>
                                        </View>
                                    ))}
                                </View>
                                {/* TABLE BODY */}

                                {reportRows.map((row, rowIndex) => (
                                    <View
                                        key={row.label}
                                        style={[
                                            styles.tableRow,
                                            rowIndex === reportRows.length - 1 &&
                                            styles.tableRowLast,
                                        ]}
                                    >
                                        {/* Row Label */}
                                        <View style={styles.cellLabel}>
                                            <Text style={styles.rowLabelText}>
                                                {row.label}
                                            </Text>
                                        </View>

                                        {/* 31 Days */}
                                        {days.map((day, dayIndex) => {
                                            const att =
                                                emp?.attendances?.[dayIndex];

                                            let value = "-";

                                            if (att) {
                                                if (row.key === "status") {
                                                    value = getStatus(att);
                                                } else if (row.isDuration) {
                                                    value = formatDuration(
                                                        att[row.key] || 0
                                                    );
                                                } else {
                                                    value =
                                                        att[row.key] || "-";
                                                }
                                            }

                                            let bgStyle = {};

                                            if (row.key === "status") {
                                                switch (value) {
                                                    case "P":
                                                        bgStyle =
                                                            styles.statusPresent;
                                                        break;

                                                    case "A":
                                                        bgStyle =
                                                            styles.statusAbsent;
                                                        break;

                                                    case "L":
                                                        bgStyle =
                                                            styles.statusLeave;
                                                        break;

                                                    case "HD":
                                                        bgStyle =
                                                            styles.statusHalfDay;
                                                        break;

                                                    case "HO":
                                                        bgStyle =
                                                            styles.statusHoliday;
                                                        break;

                                                    case "WO":
                                                        bgStyle =
                                                            styles.statusWeekOff;
                                                        break;

                                                    default:
                                                        bgStyle = {};
                                                }
                                            }

                                            return (
                                                <View
                                                    key={day}
                                                    style={[
                                                        styles.cell,
                                                        bgStyle,
                                                        dayIndex === days.length - 1 &&
                                                        styles.cellLast,
                                                    ]}
                                                >
                                                    <Text style={styles.cellText}>
                                                        {value}
                                                    </Text>
                                                </View>
                                            );
                                        })}
                                    </View>
                                ))}
                            </View>
                        </View>
                    );
                })}
            </Page>
        </Document>
    );
};

export default AttendancePDF;