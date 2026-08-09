"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

export default function Home() {
  const [files, setFiles] = useState([]);
  const [extractedData, setExtractedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [analysis, setAnalysis] = useState("");

  async function handleFiles(event) {
    const selectedFiles = Array.from(event.target.files);

    setFiles(selectedFiles);
    setExtractedData([]);
    setAnalysis("");
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

  async function analyzeResearch() {
    if (extractedData.length === 0) {
      return;
    }

    setAnalyzing(true);
    setAnalysis("");

    try {
      let transcriptText = "";

      for (const item of extractedData) {
        if (item.type === "Word" && item.text) {
          transcriptText += `\n\n--- ${item.fileName} ---\n\n`;
          transcriptText += item.text;
        }

        if (item.type === "Excel") {
          for (const sheet of item.sheets) {
            transcriptText += `\n\n--- ${item.fileName} / ${sheet.sheetName} ---\n\n`;

            for (const row of sheet.rows) {
              transcriptText += row.join(" | ") + "\n";
            }
          }
        }
      }

      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          transcript: transcriptText,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "Analysis failed.");
      }

      setAnalysis(data.result);
    } catch (error) {
      setAnalysis(`Error: ${error.message}`);
    } finally {
      setAnalyzing(false);
    }
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
        {/* Header */}

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

        {/* Upload section */}

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

          {/* Uploaded files */}

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

          {/* Loading */}

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

          {/* Analyze button */}

          {extractedData.length > 0 && !loading && (
            <button
              onClick={analyzeResearch}
              disabled={analyzing}
              style={{
                width: "100%",
                marginTop: "30px",
                padding: "16px",
                borderRadius: "10px",
                border: "none",
                background: analyzing ? "#9ca3af" : "#4f46e5",
                color: "white",
                fontSize: "16px",
                fontWeight: "bold",
                cursor: analyzing ? "not-allowed" : "pointer",
              }}
            >
              {analyzing
                ? "Analyzing Research..."
                : "Analyze Research"}
            </button>
          )}
        </div>

        {/* Extracted data */}

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

                {/* Excel */}

                {result.type === "Excel" &&
                  result.sheets.map((sheet, sheetIndex) => (
                    <div
                      key={sheetIndex}
                      style={{ marginTop: "25px" }}
                    >
                      <h4>Sheet: {sheet.sheetName}</h4>

                      <table
                        style={{
                          width: "100%",
                          borderCollapse: "collapse",
                          fontSize: "14px",
                        }}
                      >
                        <tbody>
                          {sheet.rows
                            .slice(0, 30)
                            .map((row, rowIndex) => (
                              <tr key={rowIndex}>
                                {row.map((cell, cellIndex) => (
                                  <td
                                    key={cellIndex}
                                    style={{
                                      border:
                                        "1px solid #e5e7eb",
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

                {/* Word */}

                {result.type === "Word" && (
                  <div
                    style={{
                      background: "#f9fafb",
                      padding: "20px",
                      borderRadius: "10px",
                      whiteSpace: "pre-wrap",
                      lineHeight: "1.6",
                      maxHeight: "500px",
                      overflowY: "auto",
                    }}
                  >
                    {result.text}
                  </div>
                )}

                {/* Error */}

                {result.type === "Error" && (
                  <p style={{ color: "#dc2626" }}>
                    Error: {result.message}
                  </p>
                )}
              </div>
            ))}
          </div>
        )}

        {/* AI Analysis */}

        {analyzing && (
          <div
            style={{
              marginTop: "30px",
              background: "white",
              padding: "25px",
              borderRadius: "15px",
              border: "1px solid #e5e7eb",
            }}
          >
            <h2>Analyzing research...</h2>

            <p style={{ color: "#6b7280" }}>
              ResearchOS is reviewing the transcript evidence
              and generating an initial insight.
            </p>
          </div>
        )}

        {analysis && !analyzing && (
          <div
            style={{
              marginTop: "30px",
              background: "white",
              padding: "30px",
              borderRadius: "15px",
              border: "1px solid #e5e7eb",
            }}
          >
            <h2>AI Research Analysis</h2>

            <div
              style={{
                marginTop: "20px",
                whiteSpace: "pre-wrap",
                lineHeight: "1.7",
                color: "#374151",
              }}
            >
              {analysis}
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
