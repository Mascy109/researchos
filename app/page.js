"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import pptxgen from "pptxgenjs";


export default function Home() {
  const [files, setFiles] = useState([]);
  const [extractedData, setExtractedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [analysis, setAnalysis] = useState("");
  const [analysisReady, setAnalysisReady] = useState(false);
  const [exporting, setExporting] = useState(false);

  async function handleFiles(event) {
    const selectedFiles = Array.from(event.target.files);

    setFiles(selectedFiles);
    setExtractedData([]);
    setAnalysis("");
    setAnalysisReady(false);
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
    setAnalyzing(true);
    setAnalysis("");
    setAnalysisReady(false);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          data: extractedData,
        }),
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "Analysis failed");
      }

      setAnalysis(result.analysis);
      setAnalysisReady(true);
    } catch (error) {
      setAnalysis(
        `Error: ${error.message}`
      );
    } finally {
      setAnalyzing(false);
    }
  }

  function cleanText(text) {
    if (!text) return "";

    return text
      .replace(/\*\*/g, "")
      .replace(/###/g, "")
      .replace(/##/g, "")
      .replace(/#/g, "")
      .replace(/\r/g, "")
      .trim();
  }

  function extractSections(text) {
    const cleaned = cleanText(text);

    const sections = {
      executiveSummary: "",
      findings: [],
      themes: [],
      differences: [],
      evidence: [],
      implications: [],
      recommendations: [],
    };

    const lines = cleaned
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    let currentSection = null;
    let currentFinding = null;

    for (const line of lines) {
      const lower = line.toLowerCase();

      if (
        lower.includes("executive summary") ||
        lower.includes("summary")
      ) {
        currentSection = "executiveSummary";
        continue;
      }

      if (
        lower.includes("key finding") ||
        lower.match(/^finding\s*\d+/)
      ) {
        currentSection = "findings";

        currentFinding = {
          title: line
            .replace(/key finding\s*\d*:?\s*/i, "")
            .replace(/finding\s*\d*:?\s*/i, "")
            .trim(),
          body: "",
        };

        sections.findings.push(currentFinding);
        continue;
      }

      if (
        lower.includes("themes") ||
        lower.includes("patterns")
      ) {
        currentSection = "themes";
        continue;
      }

      if (
        lower.includes("respondent differences") ||
        lower.includes("differences between respondents")
      ) {
        currentSection = "differences";
        continue;
      }

      if (
        lower.includes("evidence") ||
        lower.includes("quotes")
      ) {
        currentSection = "evidence";
        continue;
      }

      if (
        lower.includes("implications")
      ) {
        currentSection = "implications";
        continue;
      }

      if (
        lower.includes("recommendations") ||
        lower.includes("recommendation")
      ) {
        currentSection = "recommendations";
        continue;
      }

      if (currentSection === "executiveSummary") {
        sections.executiveSummary +=
          (sections.executiveSummary ? " " : "") + line;
      } else if (currentSection === "findings") {
        if (currentFinding) {
          currentFinding.body +=
            (currentFinding.body ? " " : "") + line;
        }
      } else if (currentSection === "themes") {
        sections.themes.push(line);
      } else if (currentSection === "differences") {
        sections.differences.push(line);
      } else if (currentSection === "evidence") {
        sections.evidence.push(line);
      } else if (currentSection === "implications") {
        sections.implications.push(line);
      } else if (currentSection === "recommendations") {
        sections.recommendations.push(line);
      }
    }

    return sections;
  }

  async function exportToPowerPoint() {
    if (!analysis) return;

    setExporting(true);

    try {
      const pptx = new pptxgen();

      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "ResearchOS";
      pptx.subject = "Market Research Analysis";
      pptx.title = "ResearchOS Research Analysis";
      pptx.company = "ResearchOS";
      pptx.lang = "en-US";

      pptx.theme = {
        headFontFace: "Aptos Display",
        bodyFontFace: "Aptos",
        lang: "en-US",
      };

      const sections = extractSections(analysis);

      const navy = "111827";
      const purple = "4F46E5";
      const lightPurple = "EEF2FF";
      const grey = "6B7280";
      const lightGrey = "F3F4F6";
      const white = "FFFFFF";
      const black = "111111";
      const green = "059669";

      function addHeader(slide, title, subtitle = "") {
        slide.background = {
          color: white,
        };

        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: 13.333,
          h: 0.12,
          fill: {
            color: purple,
          },
          line: {
            color: purple,
          },
        });

        slide.addText("RESEARCHOS", {
          x: 0.65,
          y: 0.45,
          w: 2,
          h: 0.3,
          fontFace: "Aptos",
          fontSize: 11,
          bold: true,
          color: purple,
          charSpacing: 1.5,
        });

        slide.addText(title, {
          x: 0.65,
          y: 0.9,
          w: 11.8,
          h: 0.65,
          fontFace: "Aptos Display",
          fontSize: 25,
          bold: true,
          color: navy,
          margin: 0,
        });

        if (subtitle) {
          slide.addText(subtitle, {
            x: 0.65,
            y: 1.55,
            w: 11.5,
            h: 0.4,
            fontSize: 11,
            color: grey,
            margin: 0,
          });
        }
      }

      function addFooter(slide, pageNumber) {
        slide.addText("ResearchOS", {
          x: 0.65,
          y: 7.15,
          w: 1.5,
          h: 0.2,
          fontSize: 8,
          color: "9CA3AF",
          margin: 0,
        });

        slide.addText(String(pageNumber), {
          x: 12.2,
          y: 7.15,
          w: 0.45,
          h: 0.2,
          fontSize: 8,
          color: "9CA3AF",
          align: "right",
          margin: 0,
        });
      }

      function addBulletList(
        slide,
        items,
        x,
        y,
        w,
        h,
        color = black
      ) {
        const validItems = items.filter(Boolean);

        if (!validItems.length) return;

        const runs = [];

        validItems.forEach((item) => {
          runs.push({
            text: cleanText(item),
            options: {
              bullet: {
                indent: 16,
              },
              hanging: 4,
              breakLine: true,
            },
          });
        });

        slide.addText(runs, {
          x,
          y,
          w,
          h,
          fontSize: 15,
          color,
          breakLine: false,
          valign: "top",
          margin: 2,
          paraSpaceAfterPt: 10,
          fit: "shrink",
        });
      }

      // ------------------------------------------------
      // SLIDE 1 — TITLE
      // ------------------------------------------------

      {
        const slide = pptx.addSlide();

        slide.background = {
          color: navy,
        };

        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: 0.18,
          h: 7.5,
          fill: {
            color: purple,
          },
          line: {
            color: purple,
          },
        });

        slide.addText("RESEARCHOS", {
          x: 0.85,
          y: 1.25,
          w: 3,
          h: 0.4,
          fontSize: 14,
          bold: true,
          color: "A5B4FC",
          charSpacing: 2,
          margin: 0,
        });

        slide.addText("Research\nAnalysis", {
          x: 0.85,
          y: 2,
          w: 8,
          h: 1.7,
          fontFace: "Aptos Display",
          fontSize: 42,
          bold: true,
          color: white,
          margin: 0,
          breakLine: false,
        });

        slide.addText(
          "From respondent conversations to structured research insights",
          {
            x: 0.9,
            y: 4.1,
            w: 7,
            h: 0.7,
            fontSize: 17,
            color: "D1D5DB",
            margin: 0,
          }
        );

        slide.addShape(pptx.ShapeType.roundRect, {
          x: 9.25,
          y: 1.7,
          w: 2.7,
          h: 2.7,
          rectRadius: 0.08,
          fill: {
            color: purple,
            transparency: 15,
          },
          line: {
            color: purple,
            transparency: 100,
          },
        });

        slide.addText("INSIGHTS", {
          x: 9.65,
          y: 2.55,
          w: 2,
          h: 0.35,
          fontSize: 13,
          bold: true,
          color: white,
          align: "center",
          charSpacing: 1.5,
          margin: 0,
        });

        slide.addText("→", {
          x: 9.95,
          y: 3.05,
          w: 1.3,
          h: 0.6,
          fontSize: 32,
          bold: true,
          color: white,
          align: "center",
          margin: 0,
        });

        slide.addText("DECISIONS", {
          x: 9.65,
          y: 3.75,
          w: 2,
          h: 0.35,
          fontSize: 13,
          bold: true,
          color: "C7D2FE",
          align: "center",
          charSpacing: 1.5,
          margin: 0,
        });

        slide.addText("Generated by ResearchOS", {
          x: 0.9,
          y: 6.7,
          w: 4,
          h: 0.3,
          fontSize: 10,
          color: "9CA3AF",
          margin: 0,
        });
      }

      // ------------------------------------------------
      // SLIDE 2 — EXECUTIVE SUMMARY
      // ------------------------------------------------

      {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Executive Summary",
          "The most important findings from the research"
        );

        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.65,
          y: 2.15,
          w: 12,
          h: 3.75,
          rectRadius: 0.05,
          fill: {
            color: lightPurple,
          },
          line: {
            color: lightPurple,
          },
        });

        slide.addText(
          cleanText(
            sections.executiveSummary ||
              "The research highlights several important themes emerging from respondent conversations."
          ),
          {
            x: 1.05,
            y: 2.7,
            w: 11.2,
            h: 2.5,
            fontSize: 21,
            bold: false,
            color: navy,
            valign: "mid",
            margin: 0,
            fit: "shrink",
            breakLine: false,
          }
        );

        addFooter(slide, 2);
      }

      // ------------------------------------------------
      // KEY FINDINGS
      // ------------------------------------------------

      const findings =
        sections.findings.length > 0
          ? sections.findings
          : [
              {
                title: "Quality drives confidence",
                body: "Respondents associate established brands with greater confidence in product quality and reliability.",
              },
              {
                title: "Trust influences choice",
                body: "Recommendations from trusted people can reduce uncertainty during purchase decisions.",
              },
              {
                title: "Price remains important",
                body: "Consumers compare prices, but lower price alone does not necessarily create preference when quality concerns exist.",
              },
            ];

      findings.slice(0, 6).forEach((finding, index) => {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          `Key Finding ${index + 1}`,
          "What the research tells us"
        );

        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.65,
          y: 2.2,
          w: 2.25,
          h: 2.25,
          rectRadius: 0.06,
          fill: {
            color: lightPurple,
          },
          line: {
            color: lightPurple,
          },
        });

        slide.addText(
          String(index + 1).padStart(2, "0"),
          {
            x: 0.65,
            y: 2.82,
            w: 2.25,
            h: 0.65,
            fontSize: 34,
            bold: true,
            color: purple,
            align: "center",
            margin: 0,
          }
        );

        slide.addText(
          cleanText(finding.title || `Key Finding ${index + 1}`),
          {
            x: 3.35,
            y: 2.25,
            w: 8.8,
            h: 1.05,
            fontFace: "Aptos Display",
            fontSize: 27,
            bold: true,
            color: navy,
            margin: 0,
            fit: "shrink",
          }
        );

        slide.addText(cleanText(finding.body), {
          x: 3.35,
          y: 3.45,
          w: 8.8,
          h: 1.5,
          fontSize: 17,
          color: grey,
          margin: 0,
          valign: "top",
          fit: "shrink",
        });

        slide.addShape(pptx.ShapeType.line, {
          x: 3.35,
          y: 5.25,
          w: 8.6,
          h: 0,
          line: {
            color: "E5E7EB",
            width: 1,
          },
        });

        slide.addText("RESEARCH IMPLICATION", {
          x: 3.35,
          y: 5.55,
          w: 3,
          h: 0.3,
          fontSize: 9,
          bold: true,
          color: purple,
          charSpacing: 1,
          margin: 0,
        });

        slide.addText(
          "This finding should be considered when evaluating consumer decision-making and brand strategy.",
          {
            x: 3.35,
            y: 5.9,
            w: 8.7,
            h: 0.65,
            fontSize: 12,
            color: black,
            margin: 0,
          }
        );

        addFooter(slide, index + 3);
      });

      // ------------------------------------------------
      // THEMES & PATTERNS
      // ------------------------------------------------

      if (sections.themes.length > 0) {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Themes & Patterns",
          "Recurring themes identified across respondent conversations"
        );

        addBulletList(
          slide,
          sections.themes.slice(0, 8),
          0.9,
          2.2,
          11.2,
          4.1
        );

        addFooter(slide, 10);
      }

      // ------------------------------------------------
      // RESPONDENT DIFFERENCES
      // ------------------------------------------------

      if (sections.differences.length > 0) {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Respondent Differences",
          "Where perspectives diverge across respondents"
        );

        addBulletList(
          slide,
          sections.differences.slice(0, 8),
          0.9,
          2.2,
          11.2,
          4.1
        );

        addFooter(slide, 11);
      }

      // ------------------------------------------------
      // EVIDENCE
      // ------------------------------------------------

      if (sections.evidence.length > 0) {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Evidence From Respondents",
          "Representative evidence supporting the findings"
        );

        const evidence = sections.evidence.slice(0, 4);

        evidence.forEach((item, index) => {
          const y = 2.15 + index * 1.15;

          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.8,
            y,
            w: 11.6,
            h: 0.85,
            rectRadius: 0.03,
            fill: {
              color: "F9FAFB",
            },
            line: {
              color: "E5E7EB",
              width: 1,
            },
          });

          slide.addText("“", {
            x: 1.05,
            y: y + 0.08,
            w: 0.35,
            h: 0.4,
            fontSize: 25,
            bold: true,
            color: purple,
            margin: 0,
          });

          slide.addText(cleanText(item), {
            x: 1.45,
            y: y + 0.17,
            w: 10.4,
            h: 0.5,
            fontSize: 12,
            italic: true,
            color: black,
            margin: 0,
            fit: "shrink",
          });
        });

        addFooter(slide, 12);
      }

      // ------------------------------------------------
      // IMPLICATIONS
      // ------------------------------------------------

      if (sections.implications.length > 0) {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Implications",
          "What these findings mean for the business"
        );

        addBulletList(
          slide,
          sections.implications.slice(0, 8),
          0.9,
          2.2,
          11.2,
          4.1
        );

        addFooter(slide, 13);
      }

      // ------------------------------------------------
      // RECOMMENDATIONS
      // ------------------------------------------------

      if (sections.recommendations.length > 0) {
        const slide = pptx.addSlide();

        addHeader(
          slide,
          "Recommendations",
          "Potential actions emerging from the research"
        );

        addBulletList(
          slide,
          sections.recommendations.slice(0, 8),
          0.9,
          2.2,
          11.2,
          4.1
        );

        addFooter(slide, 14);
      }

      // ------------------------------------------------
      // FINAL SLIDE
      // ------------------------------------------------

      {
        const slide = pptx.addSlide();

        slide.background = {
          color: navy,
        };

        slide.addText("RESEARCHOS", {
          x: 0.85,
          y: 1.45,
          w: 3,
          h: 0.4,
          fontSize: 13,
          bold: true,
          color: "A5B4FC",
          charSpacing: 2,
          margin: 0,
        });

        slide.addText("From research\nto action.", {
          x: 0.85,
          y: 2.25,
          w: 7.5,
          h: 1.5,
          fontFace: "Aptos Display",
          fontSize: 38,
          bold: true,
          color: white,
          margin: 0,
        });

        slide.addText(
          "ResearchOS turns qualitative research into structured, decision-ready insights.",
          {
            x: 0.9,
            y: 4.3,
            w: 7.2,
            h: 0.8,
            fontSize: 16,
            color: "D1D5DB",
            margin: 0,
          }
        );

        slide.addShape(pptx.ShapeType.line, {
          x: 9.1,
          y: 2.2,
          w: 2.2,
          h: 2.2,
          line: {
            color: purple,
            width: 5,
          },
        });

        slide.addText("✓", {
          x: 9.45,
          y: 2.85,
          w: 1.5,
          h: 0.8,
          fontSize: 34,
          bold: true,
          color: white,
          align: "center",
          margin: 0,
        });
      }

      await pptx.writeFile({
        fileName: "ResearchOS_Research_Analysis.pptx",
      });
    } catch (error) {
      console.error("PowerPoint export error:", error);
      alert(
        "PowerPoint export failed: " +
          (error?.message || "Unknown error")
      );
    } finally {
      setExporting(false);
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
        {/* HEADER */}

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
            Upload your qualitative research transcripts and
            transform respondent conversations into structured
            research insights.
          </p>
        </div>

        {/* UPLOAD */}

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

        {/* EXTRACTED DATA */}

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
                  result.sheets.map(
                    (sheet, sheetIndex) => (
                      <div
                        key={sheetIndex}
                        style={{ marginTop: "25px" }}
                      >
                        <h4>
                          Sheet: {sheet.sheetName}
                        </h4>

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
                              .map(
                                (row, rowIndex) => (
                                  <tr key={rowIndex}>
                                    {row.map(
                                      (
                                        cell,
                                        cellIndex
                                      ) => (
                                        <td
                                          key={
                                            cellIndex
                                          }
                                          style={{
                                            border:
                                              "1px solid #e5e7eb",
                                            padding:
                                              "8px",
                                            verticalAlign:
                                              "top",
                                          }}
                                        >
                                          {String(
                                            cell
                                          )}
                                        </td>
                                      )
                                    )}
                                  </tr>
                                )
                              )}
                          </tbody>
                        </table>

                        {sheet.rows.length >
                          30 && (
                          <p
                            style={{
                              color: "#6b7280",
                            }}
                          >
                            Showing the first 30
                            rows for preview.
                          </p>
                        )}
                      </div>
                    )
                  )}

                {result.type === "Word" && (
                  <div
                    style={{
                      background: "#f9fafb",
                      padding: "20px",
                      borderRadius: "10px",
                      whiteSpace: "pre-wrap",
                      lineHeight: "1.6",
                      color: "#374151",
                      maxHeight: "400px",
                      overflowY: "auto",
                    }}
                  >
                    {result.text}
                  </div>
                )}

                {result.type === "Error" && (
                  <p style={{ color: "#dc2626" }}>
                    Error: {result.message}
                  </p>
                )}
              </div>
            ))}

            {/* ANALYZE BUTTON */}

            {!analysisReady && (
              <div
                style={{
                  marginTop: "30px",
                  textAlign: "center",
                }}
              >
                <button
                  onClick={analyzeResearch}
                  disabled={analyzing}
                  style={{
                    background: analyzing
                      ? "#9ca3af"
                      : "#4f46e5",
                    color: "white",
                    border: "none",
                    borderRadius: "10px",
                    padding: "16px 35px",
                    fontSize: "16px",
                    fontWeight: "bold",
                    cursor: analyzing
                      ? "not-allowed"
                      : "pointer",
                  }}
                >
                  {analyzing
                    ? "Analyzing Research..."
                    : "Analyze Research →"}
                </button>
              </div>
            )}
          </div>
        )}

        {/* ANALYSIS */}

        {analysis && (
          <div
            style={{
              marginTop: "40px",
              background: "white",
              borderRadius: "18px",
              padding: "35px",
              border: "1px solid #e5e7eb",
            }}
          >
            <h2
              style={{
                marginTop: 0,
                color: "#111827",
              }}
            >
              AI Research Analysis
            </h2>

            <div
              style={{
                marginTop: "25px",
                background: "#f9fafb",
                padding: "25px",
                borderRadius: "12px",
                whiteSpace: "pre-wrap",
                lineHeight: "1.7",
                color: "#374151",
              }}
            >
              {analysis}
            </div>

            {analysisReady && (
              <div
                style={{
                  marginTop: "30px",
                  display: "flex",
                  justifyContent: "center",
                }}
              >
                <button
                  onClick={exportToPowerPoint}
                  disabled={exporting}
                  style={{
                    background: exporting
                      ? "#9ca3af"
                      : "#111827",
                    color: "white",
                    border: "none",
                    borderRadius: "10px",
                    padding: "16px 35px",
                    fontSize: "16px",
                    fontWeight: "bold",
                    cursor: exporting
                      ? "not-allowed"
                      : "pointer",
                  }}
                >
                  {exporting
                    ? "Creating PowerPoint..."
                    : "Export to PowerPoint →"}
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </main>
  );
}
