"use client";

import { useState } from "react";

export default function Home() {
  const [files, setFiles] = useState([]);

  const handleFiles = (event) => {
    const selected = Array.from(event.target.files);
    setFiles(selected);
  };

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f6f7fb",
        fontFamily: "Arial, sans-serif",
        padding: "50px 25px",
      }}
    >
      <div style={{ maxWidth: "1000px", margin: "auto" }}>

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
              maxWidth: "650px",
              lineHeight: "1.6",
            }}
          >
            Upload your qualitative research transcripts and transform
            respondent conversations into structured insights, evidence,
            and eventually a complete presentation.
          </p>
        </div>

        {/* Upload box */}

        <div
          style={{
            background: "white",
            borderRadius: "18px",
            padding: "35px",
            border: "1px solid #e5e7eb",
            boxShadow: "0 10px 30px rgba(0,0,0,0.04)",
          }}
        >
          <h2 style={{ marginTop: 0 }}>Upload research transcripts</h2>

          <p style={{ color: "#6b7280" }}>
            Upload Word interview transcripts and Excel transcript files.
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
              .DOCX and .XLSX supported
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
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "15px 18px",
                    background: "#f9fafb",
                    borderRadius: "10px",
                    marginTop: "10px",
                    border: "1px solid #e5e7eb",
                  }}
                >
                  <span>{file.name}</span>

                  <span
                    style={{
                      color: "#16a34a",
                      fontSize: "14px",
                      fontWeight: "600",
                    }}
                  >
                    Ready
                  </span>
                </div>
              ))}
            </div>
          )}

          {/* Analyze button */}

          <button
            disabled={files.length === 0}
            style={{
              width: "100%",
              marginTop: "30px",
              padding: "16px",
              borderRadius: "10px",
              border: "none",
              background:
                files.length > 0 ? "#4f46e5" : "#d1d5db",
              color: "white",
              fontSize: "16px",
              fontWeight: "bold",
              cursor:
                files.length > 0 ? "pointer" : "not-allowed",
            }}
          >
            Analyze Research
          </button>
        </div>

        {/* Future capabilities */}

        <div
          style={{
            display: "grid",
            gridTemplateColumns:
              "repeat(auto-fit, minmax(220px, 1fr))",
            gap: "20px",
            marginTop: "25px",
          }}
        >
          <Card
            number="01"
            title="Themes"
            text="Identify recurring themes and patterns across respondents."
          />

          <Card
            number="02"
            title="Insights"
            text="Turn respondent evidence into meaningful research insights."
          />

          <Card
            number="03"
            title="Presentation"
            text="Turn the final research story into a professional presentation."
          />
        </div>
      </div>
    </main>
  );
}

function Card({ number, title, text }) {
  return (
    <div
      style={{
        background: "white",
        padding: "25px",
        borderRadius: "15px",
        border: "1px solid #e5e7eb",
      }}
    >
      <div
        style={{
          color: "#4f46e5",
          fontWeight: "bold",
          marginBottom: "12px",
        }}
      >
        {number}
      </div>

      <h3 style={{ margin: "0 0 8px" }}>{title}</h3>

      <p
        style={{
          color: "#6b7280",
          lineHeight: "1.5",
          margin: 0,
        }}
      >
        {text}
      </p>
    </div>
  );
}
