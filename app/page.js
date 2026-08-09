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
  const [activeTab, setActiveTab] = useState("summary");

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

  function goToAnalysis() {
    if (extractedData.length === 0) {
      alert("Please upload at least one transcript first.");
      return;
    }

    setStep(3);
    setActiveTab("summary");
  }

  /*
   * TEMPORARY PROTOTYPE DATA
   *
   * This is intentionally mock data.
   * Later we will replace this with the actual AI analysis
   * generated from the uploaded transcripts.
   */

  const mockAnalysis = {
    summary: {
      description:
        "The research indicates that consumers evaluate brands through a combination of perceived quality, trust and practical considerations. Price remains important, but confidence in the product and recommendations from trusted people can strongly influence the final decision.",

      findings: [
        {
          title: "Quality drives confidence",
          text:
            "Respondents associate established brands with greater confidence in product quality and reliability.",
        },
        {
          title: "Trust influences choice",
          text:
            "Recommendations from contractors, family members or experienced users can reduce uncertainty during purchase decisions.",
        },
        {
          title: "Price remains important",
          text:
            "Consumers compare prices, but lower price alone does not necessarily create preference when quality concerns exist.",
        },
      ],
    },

    themes: [
      {
        number: "01",
        title: "Quality & Performance",
        description:
          "Consumers want reassurance that the product will perform reliably over time.",
        strength: "High",
      },
      {
        number: "02",
        title: "Trust & Reputation",
        description:
          "Brand reputation and recommendations help consumers feel confident about their decision.",
        strength: "High",
      },
      {
        number: "03",
        title: "Price Sensitivity",
        description:
          "Price matters particularly when consumers compare competing options.",
        strength: "Medium",
      },
      {
        number: "04",
        title: "Availability & Convenience",
        description:
          "Easy availability can remove friction from the purchase journey.",
        strength: "Medium",
      },
    ],

    needs: [
      "Confidence that the product will perform as expected",
      "A trusted and credible brand",
      "Clear information before purchase",
      "Recommendations from knowledgeable people",
      "A price that feels justified by expected quality",
    ],

    painPoints: [
      "Uncertainty about product quality",
      "Difficulty comparing competing brands",
      "Fear of making the wrong decision",
      "Conflicting recommendations",
      "Concern that a cheaper option may compromise quality",
    ],

    motivations: [
      "Desire for reliability",
      "Protecting a major investment",
      "Positive previous experience",
      "Recommendations from trusted experts",
      "Strong brand reputation",
    ],

    barriers: [
      "Higher perceived price",
      "Limited product knowledge",
      "Lack of differentiation between brands",
      "Availability issues",
      "Conflicting advice from different sources",
    ],

    segments: [
      {
        name: "Quality Seekers",
        description:
          "Consumers who prioritize reliability and performance and are willing to consider established brands.",
        characteristics: [
          "High focus on quality",
          "Risk averse",
          "Value reputation",
          "Seek reassurance before purchase",
        ],
      },
      {
        name: "Value Conscious",
        description:
          "Consumers who actively compare prices while trying to maintain an acceptable level of quality.",
        characteristics: [
          "Price sensitive",
          "Compare alternatives",
          "Look for value",
          "Open to switching brands",
        ],
      },
      {
        name: "Recommendation Led",
        description:
          "Consumers who rely heavily on trusted individuals to simplify their purchase decision.",
        characteristics: [
          "Highly influenced by experts",
          "Seek social proof",
          "Lower product knowledge",
          "Prefer low-risk choices",
        ],
      },
    ],

    evidence: [
      {
        quote:
          "I want to know that the product is reliable because I don't want problems later.",
        respondent: "R03",
        theme: "Quality & Performance",
      },
      {
        quote:
          "The contractor suggested the brand, so I felt more confident choosing it.",
        respondent: "R07",
        theme: "Trust & Reputation",
      },
      {
        quote:
          "Price is important, but I wouldn't choose something cheaper if I don't trust the quality.",
        respondent: "R11",
        theme: "Price Sensitivity",
      },
    ],

    implications: [
      "Strengthen communication around product reliability and performance.",
      "Build credibility through trusted influencers and experts.",
      "Make the value proposition clear rather than relying solely on low pricing.",
      "Use customer proof and real-world evidence to reduce purchase uncertainty.",
    ],
  };

  function renderTabContent() {
    if (activeTab === "summary") {
      return (
        <div>
          <h2>Executive Summary</h2>

          <p className="description">
            {mockAnalysis.summary.description}
          </p>

          <div className="finding-grid">
            {mockAnalysis.summary.findings.map((finding, index) => (
              <div className="finding-card" key={index}>
                <div className="eyebrow">
                  KEY FINDING {index + 1}
                </div>

                <h3>{finding.title}</h3>

                <p>{finding.text}</p>
              </div>
            ))}
          </div>
        </div>
      );
    }

    if (activeTab === "themes") {
      return (
        <div>
          <h2>Key Themes</h2>

          <div className="theme-grid">
            {mockAnalysis.themes.map((theme) => (
              <div className="theme-card" key={theme.number}>
                <div className="theme-number">
                  {theme.number}
                </div>

                <h3>{theme.title}</h3>

                <p>{theme.description}</p>

                <span
                  className={
                    theme.strength === "High"
                      ? "signal high"
                      : "signal medium"
                  }
                >
                  {theme.strength} signal
                </span>
              </div>
            ))}
          </div>
        </div>
      );
    }

    if (activeTab === "needs") {
      return (
        <div>
          <h2>Consumer Needs</h2>

          {mockAnalysis.needs.map((item, index) => (
            <div className="list-item" key={index}>
              <strong>{index + 1}</strong>
              <span>{item}</span>
            </div>
          ))}
        </div>
      );
    }

    if (activeTab === "painpoints") {
      return (
        <div>
          <h2>Pain Points</h2>

          {mockAnalysis.painPoints.map((item, index) => (
            <div className="pain-point" key={index}>
              {item}
            </div>
          ))}
        </div>
      );
    }

    if (activeTab === "motivations") {
      return (
        <div>
          <h2>Motivations</h2>

          {mockAnalysis.motivations.map((item, index) => (
            <div className="motivation" key={index}>
              {item}
            </div>
          ))}
        </div>
      );
    }

    if (activeTab === "barriers") {
      return (
        <div>
          <h2>Barriers</h2>

          {mockAnalysis.barriers.map((item, index) => (
            <div className="barrier" key={index}>
              {item}
            </div>
          ))}
        </div>
      );
    }

    if (activeTab === "segments") {
      return (
        <div>
          <h2>Respondent Segments</h2>

          <div className="segment-grid">
            {mockAnalysis.segments.map((segment) => (
              <div className="segment-card" key={segment.name}>
                <h3>{segment.name}</h3>

                <p>{segment.description}</p>

                <ul>
                  {segment.characteristics.map(
                    (characteristic, index) => (
                      <li key={index}>{characteristic}</li>
                    )
                  )}
                </ul>
              </div>
            ))}
          </div>
        </div>
      );
    }

    if (activeTab === "evidence") {
      return (
        <div>
          <h2>Evidence & Quotes</h2>

          <p className="description">
            Every research insight should eventually be supported
            by evidence from the underlying interviews.
          </p>

          {mockAnalysis.evidence.map((item, index) => (
            <div className="quote-card" key={index}>
              <p className="quote">
                "{item.quote}"
              </p>

              <div className="quote-meta">
                <strong>{item.respondent}</strong>
                <span>•</span>
                <span>{item.theme}</span>
              </div>
            </div>
          ))}
        </div>
      );
    }

    if (activeTab === "implications") {
      return (
        <div>
          <h2>Strategic Implications</h2>

          <p className="description">
            What the research could mean for the brand, product
            or business.
          </p>

          {mockAnalysis.implications.map((item, index) => (
            <div className="implication" key={index}>
              <strong>→</strong>
              <span>{item}</span>
            </div>
          ))}
        </div>
      );
    }

    return null;
  }

  return (
    <main className="page">

      {/* HEADER */}

      <header className="header">
        <div className="header-inner">
          <div className="logo">
            RESEARCHOS
          </div>

          <div className="header-label">
            AI Research Workspace
          </div>
        </div>
      </header>

      <div className="container">

        {/* PROGRESS */}

        <div className="progress">

          <div
            className={
              step >= 1
                ? "progress-step active"
                : "progress-step"
            }
          >
            01 Project Setup
          </div>

          <div className="progress-line" />

          <div
            className={
              step >= 2
                ? "progress-step active"
                : "progress-step"
            }
          >
            02 Transcripts
          </div>

          <div className="progress-line" />

          <div
            className={
              step >= 3
                ? "progress-step active"
                : "progress-step"
            }
          >
            03 Analysis
          </div>

          <div className="progress-line" />

          <div className="progress-step">
            04 Presentation
          </div>

        </div>

        {/* PROJECT SETUP */}

        {step === 1 && (
          <div>

            <div className="hero">
              <div className="eyebrow">
                RESEARCH PROJECT
              </div>

              <h1>
                Start a research project.
              </h1>

              <p>
                Tell ResearchOS what you're researching before
                uploading your transcripts.
              </p>
            </div>

            <div className="card">

              <div className="field">
                <label>
                  Project name
                </label>

                <input
                  value={project.name}
                  onChange={(event) =>
                    updateProject(
                      "name",
                      event.target.value
                    )
                  }
                  placeholder="e.g. Cement Brand Equity Study"
                />
              </div>

              <div className="field">
                <label>
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
                />
              </div>

              <div className="field">
                <label>
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
                />
              </div>

              <div className="field">
                <label>
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
                />
              </div>

              <button
                className="primary-button"
                onClick={goToUpload}
              >
                Continue to Transcripts →
              </button>

            </div>

          </div>
        )}

        {/* TRANSCRIPTS */}

        {step === 2 && (
          <div>

            <div className="hero">
              <div className="eyebrow">
                TRANSCRIPTS
              </div>

              <h1>
                Upload your transcripts.
              </h1>

              <p>
                {project.name}
              </p>
            </div>

            <div className="card">

              <label className="upload-box">

                <div className="upload-icon">
                  ↑
                </div>

                <strong>
                  Click to upload transcripts
                </strong>

                <p>
                  Word (.docx) and Excel (.xlsx)
                </p>

                <input
                  type="file"
                  multiple
                  accept=".docx,.xlsx"
                  onChange={handleFiles}
                />

              </label>

              {files.length > 0 && (
                <div className="uploaded-files">

                  <h3>
                    {files.length} file
                    {files.length !== 1 ? "s" : ""} uploaded
                  </h3>

                  {files.map((file, index) => (
                    <div
                      className="file-row"
                      key={index}
                    >
                      {file.name}
                    </div>
                  ))}

                </div>
              )}

              {loading && (
                <p className="processing">
                  Processing files...
                </p>
              )}

              {extractedData.length > 0 && !loading && (
                <div className="success">
                  <strong>
                    ✓ Transcripts processed
                  </strong>

                  <p>
                    ResearchOS has successfully extracted
                    the transcript data.
                  </p>
                </div>
              )}

              {extractedData.length > 0 && !loading && (
                <button
                  className="dark-button"
                  onClick={goToAnalysis}
                >
                  Continue to Analysis →
                </button>
              )}

            </div>

          </div>
        )}

        {/* ANALYSIS */}

        {step === 3 && (
          <div>

            <div className="hero">

              <div className="eyebrow">
                RESEARCH ANALYSIS
              </div>

              <h1>
                {project.name}
              </h1>

              <p>
                {project.objective}
              </p>

            </div>

            <div className="prototype-notice">
              <strong>Prototype mode:</strong>{" "}
              These insights are sample outputs for now.
              Later, the AI analysis engine will generate
              these insights from your actual transcripts.
            </div>

            <div className="tabs">

              {[
                ["summary", "Executive Summary"],
                ["themes", "Key Themes"],
                ["needs", "Consumer Needs"],
                ["painpoints", "Pain Points"],
                ["motivations", "Motivations"],
                ["barriers", "Barriers"],
                ["segments", "Segments"],
                ["evidence", "Evidence & Quotes"],
                ["implications", "Implications"],
              ].map(([id, label]) => (
                <button
                  key={id}
                  onClick={() => setActiveTab(id)}
                  className={
                    activeTab === id
                      ? "tab active"
                      : "tab"
                  }
                >
                  {label}
                </button>
              ))}

            </div>

            <div className="analysis-card">
              {renderTabContent()}
            </div>

            <div className="presentation-preview">

              <div className="next-label">
                NEXT
              </div>

              <h2>
                Turn these insights into a presentation
              </h2>

              <p>
                ResearchOS will transform the research analysis
                into a structured presentation containing
                insights, evidence, implications and
                recommendations.
              </p>

              <button
                className="disabled-button"
                disabled
              >
                Presentation Builder — Coming Next
              </button>

            </div>

          </div>
        )}

      </div>

      {/* STYLES */}

      <style jsx>{`

        * {
          box-sizing: border-box;
        }

        .page {
          min-height: 100vh;
          background: #f6f7fb;
          color: #111827;
          font-family: Arial, sans-serif;
        }

        .header {
          background: #ffffff;
          border-bottom: 1px solid #e5e7eb;
        }

        .header-inner {
          max-width: 1200px;
          margin: auto;
          padding: 20px 25px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .logo {
          color: #4f46e5;
          font-weight: 800;
          font-size: 18px;
          letter-spacing: 1px;
        }

        .header-label {
          color: #6b7280;
          font-size: 14px;
        }

        .container {
          max-width: 1100px;
          margin: auto;
          padding: 45px 25px 80px;
        }

        .progress {
          display: flex;
          align-items: center;
          margin-bottom: 45px;
          font-size: 13px;
        }

        .progress-step {
          color: #9ca3af;
          font-weight: 600;
          white-space: nowrap;
        }

        .progress-step.active {
          color: #4f46e5;
        }

        .progress-line {
          flex: 1;
          height: 1px;
          background: #d1d5db;
          margin: 0 15px;
        }

        .hero {
          margin-bottom: 30px;
        }

        .eyebrow {
          color: #4f46e5;
          font-size: 12px;
          font-weight: 800;
          letter-spacing: 1px;
        }

        h1 {
          font-size: 40px;
          line-height: 1.15;
          margin: 10px 0 12px;
        }

        .hero p {
          color: #6b7280;
          font-size: 17px;
          line-height: 1.6;
          max-width: 800px;
        }

        .card,
        .analysis-card {
          background: #ffffff;
          border: 1px solid #e5e7eb;
          border-radius: 18px;
          padding: 35px;
        }

        .field {
          margin-bottom: 25px;
        }

        label {
          display: block;
          font-weight: 700;
          margin-bottom: 8px;
        }

        input,
        textarea {
          width: 100%;
          padding: 14px;
          border: 1px solid #d1d5db;
          border-radius: 10px;
          font-size: 16px;
          font-family: Arial, sans-serif;
        }

        textarea {
          resize: vertical;
        }

        .primary-button,
        .dark-button {
          width: 100%;
          padding: 16px;
          border: none;
          border-radius: 10px;
          font-size: 16px;
          font-weight: 700;
          cursor: pointer;
        }

        .primary-button {
          background: #4f46e5;
          color: white;
        }

        .dark-button {
          margin-top: 25px;
          background: #111827;
          color: white;
        }

        .upload-box {
          border: 2px dashed #c7c9d9;
          border-radius: 15px;
          padding: 55px 20px;
          text-align: center;
          cursor: pointer;
        }

        .upload-box input {
          display: none;
        }

        .upload-icon {
          font-size: 40px;
          margin-bottom: 15px;
        }

        .upload-box strong {
          font-size: 18px;
        }

        .upload-box p {
          color: #9ca3af;
        }

        .uploaded-files {
          margin-top: 30px;
        }

        .file-row {
          padding: 14px;
          background: #f9fafb;
          border-radius: 10px;
          margin-top: 8px;
          border: 1px solid #e5e7eb;
        }

        .processing {
          color: #4f46e5;
          font-weight: 600;
          margin-top: 25px;
        }

        .success {
          margin-top: 30px;
          padding: 20px;
          background: #f0fdf4;
          border-radius: 12px;
          border: 1px solid #bbf7d0;
          color: #166534;
        }

        .success p {
          margin-bottom: 0;
        }

        .prototype-notice {
          background: #fffbeb;
          border: 1px solid #fde68a;
          color: #92400e;
          padding: 15px 18px;
          border-radius: 10px;
          margin-bottom: 25px;
          font-size: 14px;
          line-height: 1.5;
        }

        .tabs {
          background: #ffffff;
          border: 1px solid #e5e7eb;
          border-radius: 14px;
          padding: 8px;
          display: flex;
          gap: 5px;
          flex-wrap: wrap;
          margin-bottom: 25px;
        }

        .tab {
          padding: 10px 14px;
          border: none;
          border-radius: 8px;
          background: transparent;
          color: #4b5563;
          font-weight: 500;
          cursor: pointer;
        }

        .tab.active {
          background: #4f46e5;
          color: white;
          font-weight: 700;
        }

        .description {
          color: #4b5563;
          font-size: 17px;
          line-height: 1.7;
        }

        .finding-grid,
        .theme-grid,
        .segment-grid {
          display: grid;
          grid-template-columns:
            repeat(auto-fit, minmax(250px, 1fr));
          gap: 18px;
          margin-top: 30px;
        }

        .finding-card,
        .theme-card,
        .segment-card {
          border: 1px solid #e5e7eb;
          border-radius: 14px;
          padding: 22px;
          background: #ffffff;
        }

        .finding-card p,
        .theme-card p,
        .segment-card p {
          color: #6b7280;
          line-height: 1.6;
        }

        .theme-number {
          color: #4f46e5;
          font-size: 28px;
          font-weight: 800;
        }

        .signal {
          display: inline-block;
          padding: 6px 10px;
          border-radius: 20px;
          font-size: 12px;
          font-weight: 700;
        }

        .signal.high {
          background: #dcfce7;
          color: #166534;
        }

        .signal.medium {
          background: #fef3c7;
          color: #92400e;
        }

        .list-item {
          display: flex;
          gap: 15px;
          padding: 18px 20px;
          border-bottom: 1px solid #e5e7eb;
        }

        .list-item strong {
          color: #4f46e5;
        }

        .pain-point,
        .motivation,
        .barrier {
          padding: 18px;
          border-radius: 10px;
          margin-bottom: 12px;
        }

        .pain-point {
          background: #fff7ed;
          border: 1px solid #fed7aa;
          color: #9a3412;
        }

        .motivation {
          background: #f0fdf4;
          border: 1px solid #bbf7d0;
          color: #166534;
        }

        .barrier {
          background: #fef2f2;
          border: 1px solid #fecaca;
          color: #991b1b;
        }

        .quote-card {
          margin-top: 20px;
          padding: 25px;
          background: #f9fafb;
          border-radius: 14px;
          border-left: 4px solid #4f46e5;
        }

        .quote {
          font-size: 18px;
          line-height: 1.6;
          font-style: italic;
          margin-top: 0;
        }

        .quote-meta {
          display: flex;
          gap: 10px;
          color: #6b7280;
          font-size: 13px;
        }

        .implication {
          display: flex;
          gap: 15px;
          padding: 20px;
          background: #eef2ff;
          border-radius: 12px;
          margin-bottom: 12px;
        }

        .implication strong {
          color: #4f46e5;
        }

        .presentation-preview {
          margin-top: 30px;
          background: #111827;
          color: white;
          border-radius: 18px;
          padding: 30px;
        }

        .next-label {
          color: #a5b4fc;
          font-size: 13px;
          font-weight: 700;
          letter-spacing: 1px;
        }

        .presentation-preview p {
          color: #d1d5db;
          line-height: 1.6;
          max-width: 750px;
        }

        .disabled-button {
          padding: 14px 20px;
          background: #374151;
          color: #9ca3af;
          border: none;
          border-radius: 9px;
          font-weight: 700;
        }

        @media (max-width: 700px) {

          .progress {
            overflow-x: auto;
          }

          .progress-line {
            min-width: 20px;
          }

          h1 {
            font-size: 32px;
          }

          .card,
          .analysis-card {
            padding: 22px;
          }

          .header-label {
            display: none;
          }

        }

      `}</style>

    </main>
  );
}
