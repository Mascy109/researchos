"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

export default function Home() {
  const [step, setStep] = useState(1);

  const [project, setProject] = useState({
    name: "",
    objective: "",
    category: "",
    audience: "",
  });

  const [files, setFiles] = useState([]);
  const [extractedData, setExtractedData] = useState([]);
  const [loading, setLoading] = useState(false);

  function updateProject(field, value) {
    setProject((previous) => ({
      ...previous,
      [field]: value,
    }));
  }

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

  function goToUpload() {
    if (!project.name.trim()) {
      alert("Please enter a project name.");
      return;
    }

    if (!project.objective.trim()) {
      alert("Please enter the research objective.");
      return;
    }

    setStep(2);
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f6f7fb",
        fontFamily: "Arial, sans-serif",
        color: "#111827",
      }}
    >
      {/* TOP BAR */}

      <header
        style={{
          background: "#ffffff",
          borderBottom: "1px solid #e5e7eb",
          padding: "20px 40px",
        }}
      >
        <div
          style={{
            maxWidth: "1200px",
            margin: "auto",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <div
            style={{
              color: "#4f46e5",
              fontWeight: "800",
              fontSize: "18px",
              letterSpacing: "1px",
            }}
          >
            RESEARCHOS
          </div>

          <div
            style={{
              color: "#6b7280",
              fontSize: "14px",
            }}
          >
            AI Research Workspace
          </div>
        </div>
      </header>

      <div
        style={{
          maxWidth: "1000px",
          margin: "auto",
          padding: "55px 25px",
        }}
      >
        {/* STEP INDICATOR */}

        <div
          style={{
            display: "flex",
            alignItems: "center",
            marginBottom: "45px",
          }}
        >
          <div
            style={{
              color: step >= 1 ? "#4f46e5" : "#9ca3af",
              fontWeight: "700",
            }}
          >
            01 Project Setup
          </div>

          <div
            style={{
              flex: 1,
              height: "1px",
              background: "#d1d5db",
              margin: "0 20px",
            }}
          />

          <div
            style={{
              color: step >= 2 ? "#4f46e5" : "#9ca3af",
              fontWeight: step >= 2 ? "700" : "400",
            }}
          >
            02 Transcripts
          </div>

          <div
            style={{
              flex: 1,
              height: "1px",
              background: "#d1d5db",
              margin: "0 20px",
            }}
          />

          <div
            style={{
              color: "#9ca3af",
            }}
          >
            03 Analysis
          </div>

          <div
            style={{
              flex: 1,
              height: "1px",
              background: "#d1d5db",
              margin: "0 20px",
            }}
          />

          <div
            style={{
              color: "#9ca3af",
            }}
          >
            04 Presentation
          </div>
        </div>

        {/* STEP 1 */}

        {step === 1 && (
          <div>
            <div style={{ marginBottom: "35px" }}>
              <h1
                style={{
                  fontSize: "42px",
                  margin: "0 0 12px 0",
                }}
              >
                Start a research project.
              </h1>

              <p
                style={{
                  color: "#6b7280",
                  fontSize: "18px",
                  lineHeight: "1.6",
                  margin: 0,
                }}
              >
                Tell ResearchOS what you're researching before
                uploading your transcripts.
              </p>
            </div>

            <div
              style={{
                background: "#ffffff",
                border: "1px solid #e5e7eb",
                borderRadius: "18px",
                padding: "35px",
              }}
            >
              {/* PROJECT NAME */}

              <div style={{ marginBottom: "25px" }}>
                <label
                  style={{
                    display: "block",
                    fontWeight: "700",
                    marginBottom: "8px",
                  }}
                >
                  Project name
                </label>

                <input
                  value={project.name}
                  onChange={(event) =>
                    updateProject("name", event.target.value)
                  }
                  placeholder="e.g. Cement Brand Equity Study"
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    padding: "14px",
                    border: "1px solid #d1d5db",
                    borderRadius: "10px",
                    fontSize: "16px",
                  }}
                />
              </div>

              {/* OBJECTIVE */}

              <div style={{ marginBottom: "25px" }}>
                <label
                  style={{
                    display: "block",
                    fontWeight: "700",
                    marginBottom: "8px",
                  }}
                >
                  Research objective
                </label>

                <textarea
                  value={project.objective}
                  onChange={(event) =>
                    updateProject(
                      "objective",
                      event.target.value
                    )
                  }
                  placeholder="What do you want this research to understand?"
                  rows={5}
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    padding: "14px",
                    border: "1px solid #d1d5db",
                    borderRadius: "10px",
                    fontSize: "16px",
                    resize: "vertical",
                    fontFamily: "Arial, sans-serif",
                  }}
                />
              </div>

              {/* CATEGORY */}

              <div style={{ marginBottom: "25px" }}>
                <label
                  style={{
                    display: "block",
                    fontWeight: "700",
                    marginBottom: "8px",
                  }}
                >
                  Category / market
                </label>

                <input
                  value={project.category}
                  onChange={(event) =>
                    updateProject(
                      "category",
                      event.target.value
                    )
                  }
                  placeholder="e.g. FMCG, Automotive, Construction"
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    padding: "14px",
                    border: "1px solid #d1d5db",
                    borderRadius: "10px",
                    fontSize: "16px",
                  }}
                />
              </div>

              {/* AUDIENCE */}

              <div style={{ marginBottom: "35px" }}>
                <label
                  style={{
                    display: "block",
                    fontWeight: "700",
                    marginBottom: "8px",
                  }}
                >
                  Target audience
                </label>

                <input
                  value={project.audience}
                  onChange={(event) =>
                    updateProject(
                      "audience",
                      event.target.value
                    )
                  }
                  placeholder="e.g. Urban home builders"
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    padding: "14px",
                    border: "1px solid #d1d5db",
                    borderRadius: "10px",
                    fontSize: "16px",
                  }}
                />
              </div>

              {/* BUTTON */}

              <button
                onClick={goToUpload}
                style={{
                  width: "100%",
                  padding: "16px",
                  background: "#4f46e5",
                  color: "#ffffff",
                  border: "none",
                  borderRadius: "10px",
                  fontSize: "16px",
                  fontWeight: "700",
                  cursor: "pointer",
                }}
              >
                Continue to Transcripts →
              </button>
            </div>
          </div>
        )}

        {/* STEP 2 */}

        {step === 2 && (
          <div>
            <div style={{ marginBottom: "35px" }}>
              <h1
                style={{
                  fontSize: "38px",
                  margin: "0 0 12px 0",
                }}
              >
                Upload your transcripts.
              </h1>

              <p
                style={{
                  color: "#6b7280",
                  fontSize: "17px",
                }}
              >
                {project.name}
              </p>
            </div>

            <div
              style={{
                background: "#ffffff",
                border: "1px solid #e5e7eb",
                borderRadius: "18px",
                padding: "35px",
              }}
            >
              <label
                style={{
                  display: "block",
                  border: "2px dashed #c7c9d9",
                  borderRadius: "15px",
                  padding: "55px 20px",
                  textAlign: "center",
                  cursor: "pointer",
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

                <strong
                  style={{
                    fontSize: "18px",
                  }}
                >
                  Click to upload transcripts
                </strong>

                <p
                  style={{
                    color: "#9ca3af",
                    marginBottom: 0,
                  }}
                >
                  Word (.docx) and Excel (.xlsx)
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
                  <h3>
                    {files.length} file
                    {files.length !== 1 ? "s" : ""} uploaded
                  </h3>

                  {files.map((file, index) => (
                    <div
                      key={index}
                      style={{
                        padding: "14px",
                        background: "#f9fafb",
                        borderRadius: "10px",
                        marginTop: "8px",
                        border:
                          "1px solid #e5e7eb",
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
                  Processing files...
                </p>
              )}

              {extractedData.length > 0 &&
                !loading && (
                  <div
                    style={{
                      marginTop: "30px",
                      padding: "20px",
                      background: "#f0fdf4",
                      borderRadius: "12px",
                      border:
                        "1px solid #bbf7d0",
                    }}
                  >
                    <strong
                      style={{
                        color: "#166534",
                      }}
                    >
                      ✓ Transcripts processed
                    </strong>

                    <p
                      style={{
                        color: "#166534",
                        marginBottom: 0,
                      }}
                    >
                      ResearchOS has successfully
                      extracted the transcript data.
                    </p>
                  </div>
                )}
            </div>

            {/* PROJECT SUMMARY */}

            <div
              style={{
                marginTop: "25px",
                background: "#ffffff",
                border: "1px solid #e5e7eb",
                borderRadius: "18px",
                padding: "30px",
              }}
            >
              <h3
                style={{
                  marginTop: 0,
                }}
              >
                Research brief
              </h3>

              <p>
                <strong>Objective:</strong>{" "}
                {project.objective}
              </p>

              {project.category && (
                <p>
                  <strong>Category:</strong>{" "}
                  {project.category}
                </p>
              )}

              {project.audience && (
                <p>
                  <strong>Audience:</strong>{" "}
                  {project.audience}
                </p>
              )}
            </div>

            {/* FUTURE WORKSPACE PREVIEW */}

            {extractedData.length > 0 &&
              !loading && (
                <div
                  style={{
                    marginTop: "30px",
                    background: "#111827",
                    color: "#ffffff",
                    borderRadius: "18px",
                    padding: "30px",
                  }}
                >
                  <div
                    style={{
                      color: "#a5b4fc",
                      fontSize: "13px",
                      fontWeight: "700",
                      letterSpacing: "1px",
                    }}
                  >
                    NEXT
                  </div>

                  <h2
                    style={{
                      margin: "10px 0",
                    }}
                  >
                    Build your research analysis
                  </h2>

                  <p
                    style={{
                      color: "#d1d5db",
                      lineHeight: "1.6",
                    }}
                  >
                    Once the analysis engine is connected,
                    ResearchOS will identify themes,
                    patterns, respondent differences,
                    evidence and insights from these
                    transcripts.
                  </p>

                  <button
                    disabled
                    style={{
                      marginTop: "10px",
                      padding: "14px 20px",
                      border: "none",
                      borderRadius: "9px",
                      background: "#374151",
                      color: "#9ca3af",
                      fontWeight: "700",
                      cursor: "not-allowed",
                    }}
                  >
                    Analysis Engine — Coming Next
                  </button>
                </div>
              )}
          </div>
        )}
      </div>
    </main>
  );
}
