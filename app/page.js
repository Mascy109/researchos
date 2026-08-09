"use client";

import { useState } from "react";

export default function Home() {
  const [files, setFiles] = useState([]);

  function handleFiles(event) {
    const selectedFiles = Array.from(event.target.files);
    setFiles(selectedFiles);
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f7f8fc",
        fontFamily: "Arial, sans-serif",
        color: "#172033",
        padding: "40px",
      }}
    >
      <div
        style={{
          maxWidth: "1000px",
          margin: "0 auto",
        }}
      >
        <header style={{ marginBottom: "40px" }}>
          <div
            style={{
              fontSize: "14px",
              fontWeight: "700",
              color: "#6366f1",
              marginBottom: "10px",
            }}
          >
            RESEARCHOS
          </div>

          <h1
            style={{
              fontSize: "42px",
              margin: "0 0 12px",
            }}
          >
            Turn transcripts into research insights.
          </h1>

          <p
            style={{
              fontSize: "18px",
              color: "#667085",
              maxWidth: "650px",
              lineHeight: "1.6",
            }}
          >
            Upload your qualitative research transcripts and use AI to
            identify themes, insights, evidence and opportunities.
          </p>
        </header>

        <section
          style={{
            background: "white",
            borderRadius: "20px",
            padding: "35px",
            border: "1px solid #e4e7ec",
            boxShadow: "0 8px 30px rgba(0,0,0,0.04)",
          }}
        >
          <h2 style={{ marginTop: 0 }}>Upload transcripts</h2>

          <p style={{ color: "#667085" }}>
            Upload Word (.docx) and Excel (.xlsx) interview transcripts.
          </p>

          <label
            style={{
              display: "block",
              border: "2px dashed #c7c9d9",
              borderRadius: "16px",
              padding: "55px 20px",
              textAlign: "center",
              cursor: "pointer",
              marginTop: "25px",
            }}
          >
            <div style={{ fontSize: "42px", marginBottom: "15px" }}>↑</div>

            <strong style={{ fontSize: "18px" }}>
              Click to upload your transcripts
            </strong>

            <div
              style={{
                marginTop: "8px",
                color: "#667085",
              }}
            >
              Word and Excel files supported
            </div>

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
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "15px 18px",
                    background: "#f7f8fc",
                    borderRadius: "10px",
                    marginTop: "10px",
                  }}
                >
                  <span>{file.name}</span>

                  <span
                    style={{
                      fontSize: "13px",
                      color: "#667085",
                    }}
                  >
                    Ready
                  </span>
                </div>
              ))}
            </div>
          )}

          <button
            style={{
              width: "100%",
              marginTop: "30px",
              padding: "16px",
              borderRadius: "10px",
              border: "none",
              background: files.length ? "#4f46e5" : "#d0d5dd",
              color: "white",
              fontSize: "16px",
              fontWeight: "700",
              cursor: files.length ? "pointer" : "not-allowed",
            }}
            disabled={!files.length}
          >
            Analyze Research
          </button>
        </section>

        <section
          style={{
            marginTop: "25px",
            display: "grid",
            gridTemplateColumns: "repeat(3, 1fr)",
            gap: "20px",
          }}
        >
          {[
            ["01", "Themes", "Identify recurring themes and patterns."],
            ["02", "Insights", "Turn evidence into meaningful research insights."],
            ["03", "Presentation", "Turn the findings into a presentation."],
          ].map(([number, title, description]) => (
            <div
              key={number}
              style={{
                background: "white",
                padding: "25px",
                borderRadius: "16px",
                border: "1px solid #e4e7ec",
              }}
            >
              <div
                style={{
                  color: "#6366f1",
                  fontWeight: "700",
                  marginBottom: "12px",
                }}
              >
                {number}
              </div>

              <h3 style={{ margin: "0 0 8px" }}>{title}</h3>

              <p
                style={{
                  color: "#667085",
                  lineHeight: "1.5",
                  margin: 0,
                }}
              >
                {description}
              </p>
            </div>
          ))}
        </section>
      </div>
    </main>
  );
}
