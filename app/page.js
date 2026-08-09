"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

export default function Home() {
  const [files, setFiles] = useState([]);
  const [extractedData, setExtractedData] = useState([]);
  const [loading, setLoading] = useState(false);

  async function handleFiles(event) {
    const selectedFiles = Array.from(event.target.files);

    setFiles(selectedFiles);
    setExtractedData([]);
    setLoading(true);

    const results = [];

    for (const file of selectedFiles) {
      try {
        if (file.name.toLowerCase().endsWith(".xlsx")) {
          const buffer = await file.arrayBuffer();

          const workbook = XLSX.read(buffer, {
            type: "array",
          });

          const sheets = workbook.SheetNames.map((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];

            return {
              sheetName,
              rows: XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: "",
              }),
            };
          });

          results.push({
            fileName: file.name,
            type: "Excel",
            sheets,
          });
        }

       if (file.name.toLowerCase().endsWith(".docx")) {
  const buffer = await file.arrayBuffer();

  const result = await mammoth.extractRawText({
    arrayBuffer: buffer,
  });

  results.push({
    fileName: file.name,
    type: "Word",
    text: result.value,
  });
}
      } catch (error) {
        results.push({
          fileName: file.name,
          type: "Error",
          message: error.message,
        });
      }
    }

    setExtractedData(results);
    setLoading(false);
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f6f7fb",
        fontFamily: "Arial, sans-serif",
        padding: "50px 25px",
      }}
    >
      <div
        style={{
          maxWidth: "1100px",
          margin: "auto",
        }}
      >
        <div style={{ marginBottom: "40px" }}>
          <div
            style={{
              color: "#4f46e5",
              fontWeight: "bold",
              fontSize: "14px",
              letterSpacing: "1px",
            }}
          >
            RESEARCHOS
          </div>

          <h1
            style={{
              fontSize: "42px",
              margin: "12px 0",
              color: "#111827",
            }}
          >
            Turn transcripts into research insights.
          </h1>

          <p
            style={{
              color: "#6b7280",
              fontSize: "18px",
              maxWidth: "700px",
              lineHeight: "1.6",
            }}
          >
            Upload your qualitative research transcripts and transform
            respondent conversations into structured research insights.
          </p>
        </div>

        <div
          style={{
            background: "white",
            borderRadius: "18px",
            padding: "35px",
            border: "1px solid #e5e7eb",
          }}
        >
          <h2 style={{ marginTop: 0 }}>
            Upload research transcripts
          </h2>

          <p style={{ color: "#6b7280" }}>
            Supported formats: Word (.docx) and Excel (.xlsx)
          </p>

          <label
            style={{
              display: "block",
              border: "2px dashed #c7c9d9",
              borderRadius: "15px",
              padding: "55px 20px",
              textAlign: "center",
              cursor: "pointer",
              marginTop: "25px",
            }}
          >
            <div
              style={{
                fontSize: "40px",
                marginBottom: "15px",
              }}
            >
              ↑
            </div>

            <strong style={{ fontSize: "18px" }}>
              Click to upload files
            </strong>

            <p
              style={{
                color: "#9ca3af",
                marginBottom: 0,
              }}
            >
              You can select multiple files
            </p>

            <input
              type="file"
              multiple
              accept=".docx,.xlsx"
              onChange={handleFiles}
              style={{ display: "none" }}
            />
          </label>

          {files.length > 0 && (
            <div style={{ marginTop: "30px" }}>
              <h3>Uploaded files</h3>

              {files.map((file, index) => (
                <div
                  key={index}
                  style={{
                    padding: "15px",
                    background: "#f9fafb",
                    borderRadius: "10px",
                    marginTop: "10px",
                    border: "1px solid #e5e7eb",
                  }}
                >
                  {file.name}
                </div>
              ))}
            </div>
          )}

          {loading && (
            <p
              style={{
                marginTop: "25px",
                color: "#4f46e5",
                fontWeight: "600",
              }}
            >
              Reading files...
            </p>
          )}
        </div>

        {extractedData.length > 0 && (
          <div style={{ marginTop: "30px" }}>
            <h2>Extracted Research Data</h2>

            {extractedData.map((result, index) => (
              <div
                key={index}
                style={{
                  background: "white",
                  borderRadius: "15px",
                  padding: "25px",
                  marginTop: "20px",
                  border: "1px solid #e5e7eb",
                  overflowX: "auto",
                }}
              >
                <h3>{result.fileName}</h3>

                {result.type === "Excel" &&
                  result.sheets.map((sheet, sheetIndex) => (
                    <div key={sheetIndex} style={{ marginTop: "25px" }}>
                      <h4>Sheet: {sheet.sheetName}</h4>

                      <table
                        style={{
                          width: "100%",
                          borderCollapse: "collapse",
                          fontSize: "14px",
                        }}
                      >
                        <tbody>
                          {sheet.rows.slice(0, 30).map((row, rowIndex) => (
                            <tr key={rowIndex}>
                              {row.map((cell, cellIndex) => (
                                <td
                                  key={cellIndex}
                                  style={{
                                    border: "1px solid #e5e7eb",
                                    padding: "8px",
                                    verticalAlign: "top",
                                  }}
                                >
                                  {String(cell)}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>

                      {sheet.rows.length > 30 && (
                        <p style={{ color: "#6b7280" }}>
                          Showing the first 30 rows for preview.
                        </p>
                      )}
                    </div>
                  ))}

                {result.type === "Word" && (
                  <p style={{ color: "#6b7280" }}>
                    {result.message}
                  </p>
                )}

                {result.type === "Error" && (
                  <p style={{ color: "#dc2626" }}>
                    Error: {result.message}
                  </p>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </main>
  );
}
