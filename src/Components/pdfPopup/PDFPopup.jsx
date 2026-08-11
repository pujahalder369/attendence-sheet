import { useEffect, useState } from "react";
import { pdf } from "@react-pdf/renderer";
import { saveAs } from "file-saver";
import toast, { Toaster } from "react-hot-toast";
import { downloadExcel } from "../excelPopup/ExcelPopup";
import AttendancePDF from "../AttendencePDF/AttendancePDF";

const PDFPopup = ({ closePopup }) => {
  const [attendanceData, setAttendanceData] = useState({
    results: [],
  });

  const [selectedEmployees, setSelectedEmployees] = useState([]);
  const [exportType, setExportType] = useState("pdf");
  const [isGenerating, setIsGenerating] = useState(false);

  useEffect(() => {
    fetch("/data.json")
      .then((res) => res.json())
      .then((json) => setAttendanceData(json))
      .catch((err) => {
        console.error(err);
        toast.error("Failed to load attendance data");
      });
  }, []);

  // SELECT EMPLOYEE
  const handleCheckboxChange = (id) => {
    setSelectedEmployees((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id],
    );
  };

  // PDF
  const downloadPDF = async (selectedData) => {
    const blob = await pdf(
      <AttendancePDF data={selectedData} selectedIds={selectedEmployees} />,
    ).toBlob();

    saveAs(blob, "attendance-report.pdf");
  };

  // MAIN DOWNLOAD
  const handleDownload = async () => {
    const selectedData = attendanceData.results.filter((entry) =>
      selectedEmployees.includes(entry?.employee?.id),
    );

    if (selectedData.length === 0) {
      toast.error("Please select at least one employee.");
      return;
    }

    setIsGenerating(true);

    const loading = toast.loading(
      exportType === "pdf" ? "Generating PDF..." : "Generating Excel...",
    );

    try {
      if (exportType === "pdf") {
        await downloadPDF(selectedData);

        toast.dismiss(loading);
        toast.success("PDF downloaded successfully!");
      } else {
        downloadExcel(selectedData);

        toast.dismiss(loading);
        toast.success("Excel downloaded successfully!");
      }
    } catch (error) {
      console.error(error);

      toast.dismiss(loading);

      toast.error(
        exportType === "pdf"
          ? "Failed to generate PDF."
          : "Failed to generate Excel.",
      );
    } finally {
      setIsGenerating(false);
    }
  };

  const handleSelectAll = (check) => {
    if (check) {
      const allEmployeeIds = 
      attendanceData.results.map((entry) => entry?.employee?.id).filter(Boolean);
      setSelectedEmployees(allEmployeeIds);
    }else {
      setSelectedEmployees([]);
    }
  }

  return (
    <>
      <Toaster position="top-right" />

      <div className="fixed inset-0 bg-black/50 flex justify-center items-center">
        <div className="bg-white p-6 rounded-lg w-[500px]">
          {/* =========================
              EMPLOYEES + EXPORT
          ========================= */}
          <div className="mb-5 grid grid-cols-2 gap-4">
            {/* EMPLOYEES */}
            <div>
              <h2 className="font-semibold text-lg mb-3">Employees</h2>

              <div className="left_box flex flex-col gap-2">
                {attendanceData.results.map((entry) => {
                  const emp = entry?.employee;
                  const employeeId = emp?.id;

                  return (
                    <label
                      key={employeeId}
                      htmlFor={`emp-${employeeId}`}
                      className="flex items-center gap-2 cursor-pointer"
                    >
                      <input
                        type="checkbox"
                        id={`emp-${employeeId}`}
                        checked={selectedEmployees.includes(employeeId)}
                        onChange={() => handleCheckboxChange(employeeId)}
                        className="w-4 h-4"
                      />

                      <span>
                        {emp?.first_name} {emp?.last_name}
                      </span>
                    </label>
                  );
                })}

                <label
                  htmlFor="select-all"
                  className="flex items-center gap-2 cursor-pointer mb-3"
                >
                  <input
                    type="checkbox"
                    id="select-all"
                    checked={
                      attendanceData.results.length > 0 &&
                      selectedEmployees.length === attendanceData.results.length
                    }
                    onChange={(e) => handleSelectAll(e.target.checked)}
                    className="w-4 h-4"
                  />

                  <span>Select All</span>
                </label>
              </div>
            </div>

            {/* EXPORT TYPE */}
            <div>
              <h2 className="font-semibold text-lg mb-3">Export In</h2>

              <div className="left_box flex flex-col gap-3">
                {/* PDF */}
                <label
                  htmlFor="pdf"
                  className="flex items-center gap-2 cursor-pointer"
                >
                  <input
                    type="radio"
                    name="export"
                    id="pdf"
                    value="pdf"
                    checked={exportType === "pdf"}
                    onChange={() => setExportType("pdf")}
                    className="h-4 w-4"
                  />

                  <span>PDF</span>
                </label>

                {/* EXCEL */}
                <label
                  htmlFor="excel"
                  className="flex items-center gap-2 cursor-pointer"
                >
                  <input
                    type="radio"
                    name="export"
                    id="excel"
                    value="excel"
                    checked={exportType === "excel"}
                    onChange={() => setExportType("excel")}
                    className="h-4 w-4"
                  />

                  <span>Excel</span>
                </label>
              </div>
            </div>
          </div>

          {/* =========================
              BUTTONS
          ========================= */}
          <div className="flex flex-col justify-center items-center gap-3">
            <button
              onClick={handleDownload}
              disabled={isGenerating}
              className="bg-[#EDEB62] p-2 rounded-lg flex w-full justify-center font-semibold disabled:opacity-70"
            >
              {isGenerating
                ? "Generating..."
                : `Download ${exportType === "pdf" ? "PDF" : "Excel"}`}
            </button>

            <button
              onClick={closePopup}
              disabled={isGenerating}
              className="bg-[#EDEB62] p-2 rounded-lg flex w-full justify-center font-semibold"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </>
  );
};

export default PDFPopup;
