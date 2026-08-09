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

  const [presentationType, setPresentationType] =
    useState("research");

  const [slideCount, setSlideCount] = useState("8");

  const [presentationStyle, setPresentationStyle] =
    useState("consulting");

  const [presentationGenerated, setPresentationGenerated] =
    useState(false);

  const [currentSlide, setCurrentSlide] = useState(0);

  const [slides, setSlides] = useState([]);

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

  function goToPresentation() {
    setStep(4);
    setPresentationGenerated(false);
  }

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

  function createSlides() {
    let generatedSlides = [];

    generatedSlides.push({
      type: "title",
      title:
        project.name || "Research Study",
      subtitle:
        "Consumer Research Findings & Strategic Implications",
      label: "RESEARCHOS",
    });

    generatedSlides.push({
      type: "executive",
      title: "Executive Summary",
      subtitle:
        "What the research tells us",
      bullets: [
        "Quality and reliability are major drivers of consumer confidence.",
        "Trust and recommendations reduce uncertainty during purchase decisions.",
        "Price remains important, but consumers do not always choose the cheapest option.",
      ],
    });

    generatedSlides.push({
      type: "themes",
      title: "Four themes shape the decision",
      subtitle:
        "The strongest patterns emerging from the research",
      themes: [
        ["01", "Quality & Performance", "HIGH"],
        ["02", "Trust & Reputation", "HIGH"],
        ["03", "Price Sensitivity", "MEDIUM"],
        ["04", "Availability & Convenience", "MEDIUM"],
      ],
    });

    generatedSlides.push({
      type: "needs",
      title: "Consumers need confidence before they commit",
      subtitle:
        "The underlying needs expressed across the research",
      bullets: mockAnalysis.needs,
    });

    generatedSlides.push({
      type: "painpoints",
      title: "Uncertainty remains a major barrier",
      subtitle:
        "Where consumers experience friction",
      bullets: mockAnalysis.painPoints,
    });

    generatedSlides.push({
      type: "segments",
      title: "Three consumer mindsets emerge",
      subtitle:
        "Different consumers approach the decision differently",
      segments: mockAnalysis.segments,
    });

    generatedSlides.push({
      type: "evidence",
      title: "The findings are grounded in consumer voices",
      subtitle:
        "Illustrative evidence from the research",
      quotes: mockAnalysis.evidence,
    });

    generatedSlides.push({
      type: "implications",
      title: "What this means for the brand",
      subtitle:
        "Strategic implications from the research",
      bullets: mockAnalysis.implications,
    });

    if (slideCount === "10") {
      generatedSlides.splice(7, 0, {
        type: "motivations",
        title: "What motivates consumers",
        subtitle:
          "Key drivers behind positive decision-making",
        bullets: mockAnalysis.motivations,
      });

      generatedSlides.splice(8, 0, {
        type: "barriers",
        title: "What can prevent conversion",
        subtitle:
          "Barriers that need to be addressed",
        bullets: mockAnalysis.barriers,
      });
    }

    if (slideCount === "6") {
      generatedSlides = generatedSlides.filter(
        (_, index) =>
          ![3, 4].includes(index)
      );
    }

    setSlides(generatedSlides);
    setCurrentSlide(0);
    setPresentationGenerated(true);
  }

  function updateSlide(field, value) {
    setSlides((previous) =>
      previous.map((slide, index) =>
        index === currentSlide
          ? {
              ...slide,
              [field]: value,
            }
          : slide
      )
    );
  }

  function renderTabContent() {
    if (activeTab === "summary") {
      return (
        <div>
          <h2>Executive Summary</h2>

          <p className="description">
            {mockAnalysis.summary.description}
          </p>

          <div className="finding-grid">
            {mockAnalysis.summary.findings.map(
              (finding, index) => (
                <div
                  className="finding-card"
                  key={index}
                >
                  <div className="eyebrow">
                    KEY FINDING {index + 1}
                  </div>

                  <h3>{finding.title}</h3>

                  <p>{finding.text}</p>
                </div>
              )
            )}
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
              <div
                className="theme-card"
                key={theme.number}
              >
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

          {mockAnalysis.needs.map(
            (item, index) => (
              <div
                className="list-item"
                key={index}
              >
                <strong>{index + 1}</strong>
                <span>{item}</span>
              </div>
            )
          )}
        </div>
      );
    }

    if (activeTab === "painpoints") {
      return (
        <div>
          <h2>Pain Points</h2>

          {mockAnalysis.painPoints.map(
            (item, index) => (
              <div
                className="pain-point"
                key={index}
              >
                {item}
              </div>
            )
          )}
        </div>
      );
    }

    if (activeTab === "motivations") {
      return (
        <div>
          <h2>Motivations</h2>

          {mockAnalysis.motivations.map(
            (item, index) => (
              <div
                className="motivation"
                key={index}
              >
                {item}
              </div>
            )
          )}
        </div>
      );
    }

    if (activeTab === "barriers") {
      return (
        <div>
          <h2>Barriers</h2>

          {mockAnalysis.barriers.map(
            (item, index) => (
              <div
                className="barrier"
                key={index}
              >
                {item}
              </div>
            )
          )}
        </div>
      );
    }

    if (activeTab === "segments") {
      return (
        <div>
          <h2>Respondent Segments</h2>

          <div className="segment-grid">
            {mockAnalysis.segments.map(
              (segment) => (
                <div
                  className="segment-card"
                  key={segment.name}
                >
                  <h3>{segment.name}</h3>

                  <p>{segment.description}</p>

                  <ul>
                    {segment.characteristics.map(
                      (characteristic, index) => (
                        <li key={index}>
                          {characteristic}
                        </li>
                      )
                    )}
                  </ul>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    if (activeTab === "evidence") {
      return (
        <div>
          <h2>Evidence & Quotes</h2>

          <p className="description">
            Every research insight should eventually
            be supported by evidence from the
            underlying interviews.
          </p>

          {mockAnalysis.evidence.map(
            (item, index) => (
              <div
                className="quote-card"
                key={index}
              >
                <p className="quote">
                  "{item.quote}"
                </p>

                <div className="quote-meta">
                  <strong>
                    {item.respondent}
                  </strong>

                  <span>•</span>

                  <span>{item.theme}</span>
                </div>
              </div>
            )
          )}
        </div>
      );
    }

    if (activeTab === "implications") {
      return (
        <div>
          <h2>Strategic Implications</h2>

          <p className="description">
            What the research could mean for the
            brand, product or business.
          </p>

          {mockAnalysis.implications.map(
            (item, index) => (
              <div
                className="implication"
                key={index}
              >
                <strong>→</strong>

                <span>{item}</span>
              </div>
            )
          )}
        </div>
      );
    }

    return null;
  }

  function renderPresentationSlide(slide) {
    if (slide.type === "title") {
      return (
        <div className="ppt-title-slide">
          <div className="ppt-brand">
            {slide.label}
          </div>

          <h1>{slide.title}</h1>

          <p>{slide.subtitle}</p>

          <div className="ppt-footer">
            Research analysis • {project.category || "Market Research"}
          </div>
        </div>
      );
    }

    if (slide.type === "executive") {
      return (
        <div className="ppt-content-slide">
          <div className="ppt-small-label">
            EXECUTIVE SUMMARY
          </div>

          <h1>{slide.title}</h1>

          <p className="ppt-subtitle">
            {slide.subtitle}
          </p>

          <div className="ppt-bullets">
            {slide.bullets.map(
              (bullet, index) => (
                <div
                  className="ppt-bullet"
                  key={index}
                >
                  <div className="ppt-bullet-number">
                    {String(index + 1).padStart(
                      2,
                      "0"
                    )}
                  </div>

                  <div>{bullet}</div>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    if (slide.type === "themes") {
      return (
        <div className="ppt-content-slide">
          <div className="ppt-small-label">
            KEY THEMES
          </div>

          <h1>{slide.title}</h1>

          <p className="ppt-subtitle">
            {slide.subtitle}
          </p>

          <div className="ppt-theme-grid">
            {slide.themes.map(
              (theme) => (
                <div
                  className="ppt-theme"
                  key={theme[0]}
                >
                  <div className="ppt-theme-number">
                    {theme[0]}
                  </div>

                  <h3>{theme[1]}</h3>

                  <span>
                    {theme[2]} SIGNAL
                  </span>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    if (
      slide.type === "needs" ||
      slide.type === "painpoints" ||
      slide.type === "motivations" ||
      slide.type === "barriers" ||
      slide.type === "implications"
    ) {
      return (
        <div className="ppt-content-slide">
          <div className="ppt-small-label">
            RESEARCH INSIGHT
          </div>

          <h1>{slide.title}</h1>

          <p className="ppt-subtitle">
            {slide.subtitle}
          </p>

          <div className="ppt-bullets">
            {slide.bullets.map(
              (bullet, index) => (
                <div
                  className="ppt-bullet"
                  key={index}
                >
                  <div className="ppt-bullet-dot">
                    ●
                  </div>

                  <div>{bullet}</div>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    if (slide.type === "segments") {
      return (
        <div className="ppt-content-slide">
          <div className="ppt-small-label">
            RESPONDENT SEGMENTS
          </div>

          <h1>{slide.title}</h1>

          <p className="ppt-subtitle">
            {slide.subtitle}
          </p>

          <div className="ppt-segment-grid">
            {slide.segments.map(
              (segment) => (
                <div
                  className="ppt-segment"
                  key={segment.name}
                >
                  <h3>{segment.name}</h3>

                  <p>
                    {segment.description}
                  </p>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    if (slide.type === "evidence") {
      return (
        <div className="ppt-content-slide">
          <div className="ppt-small-label">
            CONSUMER VOICE
          </div>

          <h1>{slide.title}</h1>

          <p className="ppt-subtitle">
            {slide.subtitle}
          </p>

          <div className="ppt-quote-grid">
            {slide.quotes.map(
              (quote, index) => (
                <div
                  className="ppt-quote"
                  key={index}
                >
                  <div className="quote-mark">
                    "
                  </div>

                  <p>{quote.quote}</p>

                  <small>
                    {quote.respondent} •{" "}
                    {quote.theme}
                  </small>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    return null;
  }

  return (
    <main className="page">

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

          <div
            className={
              step >= 4
                ? "progress-step active"
                : "progress-step"
            }
          >
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
                Tell ResearchOS what you're researching
                before uploading your transcripts.
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
                    {files.length !== 1
                      ? "s"
                      : ""}{" "}
                    uploaded
                  </h3>

                  {files.map(
                    (file, index) => (
                      <div
                        className="file-row"
                        key={index}
                      >
                        {file.name}
                      </div>
                    )
                  )}

                </div>
              )}

              {loading && (
                <p className="processing">
                  Processing files...
                </p>
              )}

              {extractedData.length > 0 &&
                !loading && (
                  <div className="success">

                    <strong>
                      ✓ Transcripts processed
                    </strong>

                    <p>
                      ResearchOS has successfully
                      extracted the transcript data.
                    </p>

                  </div>
                )}

              {extractedData.length > 0 &&
                !loading && (
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

              <strong>
                Prototype mode:
              </strong>{" "}
              These insights are sample outputs
              for now. Later, the AI analysis engine
              will generate these insights from your
              actual transcripts.

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
              ].map(
                ([id, label]) => (
                  <button
                    key={id}
                    onClick={() =>
                      setActiveTab(id)
                    }
                    className={
                      activeTab === id
                        ? "tab active"
                        : "tab"
                    }
                  >
                    {label}
                  </button>
                )
              )}

            </div>

            <div className="analysis-card">
              {renderTabContent()}
            </div>

            <div className="presentation-preview">

              <div className="next-label">
                NEXT
              </div>

              <h2>
                Turn these insights into a
                presentation
              </h2>

              <p>
                ResearchOS will transform the
                research analysis into a structured
                presentation containing insights,
                evidence, implications and
                recommendations.
              </p>

              <button
                className="presentation-button"
                onClick={goToPresentation}
              >
                Build Presentation →
              </button>

            </div>

          </div>
        )}

        {/* PRESENTATION BUILDER */}

        {step === 4 && !presentationGenerated && (
          <div>

            <div className="hero">

              <div className="eyebrow">
                PRESENTATION BUILDER
              </div>

              <h1>
                Build your research deck.
              </h1>

              <p>
                Turn the analysis into a structured
                presentation ready for clients,
                stakeholders or internal teams.
              </p>

            </div>

            <div className="builder-layout">

              <div className="card">

                <h2>
                  Presentation setup
                </h2>

                <p className="builder-description">
                  Choose how ResearchOS should structure
                  the presentation.
                </p>

                <div className="field">

                  <label>
                    Presentation type
                  </label>

                  <select
                    value={presentationType}
                    onChange={(event) =>
                      setPresentationType(
                        event.target.value
                      )
                    }
                  >
                    <option value="research">
                      Research Findings
                    </option>

                    <option value="executive">
                      Executive Summary
                    </option>

                    <option value="client">
                      Client Presentation
                    </option>
                  </select>

                </div>

                <div className="field">

                  <label>
                    Number of slides
                  </label>

                  <select
                    value={slideCount}
                    onChange={(event) =>
                      setSlideCount(
                        event.target.value
                      )
                    }
                  >
                    <option value="6">
                      6 slides
                    </option>

                    <option value="8">
                      8 slides
                    </option>

                    <option value="10">
                      10 slides
                    </option>
                  </select>

                </div>

                <div className="field">

                  <label>
                    Presentation style
                  </label>

                  <select
                    value={presentationStyle}
                    onChange={(event) =>
                      setPresentationStyle(
                        event.target.value
                      )
                    }
                  >
                    <option value="consulting">
                      Consulting
                    </option>

                    <option value="minimal">
                      Minimal
                    </option>

                    <option value="corporate">
                      Corporate
                    </option>
                  </select>

                </div>

                <button
                  className="primary-button"
                  onClick={createSlides}
                >
                  Generate Presentation →
                </button>

              </div>

              <div className="builder-preview">

                <div className="preview-label">
                  DECK PREVIEW
                </div>

                <div className="mini-slide">

                  <div className="mini-brand">
                    RESEARCHOS
                  </div>

                  <div className="mini-title">
                    {project.name ||
                      "Research Study"}
                  </div>

                  <div className="mini-subtitle">
                    Consumer Research Findings
                    & Strategic Implications
                  </div>

                  <div className="mini-lines">
                    <span />
                    <span />
                    <span />
                  </div>

                </div>

                <p>
                  Your presentation will be
                  generated from the research
                  analysis.
                </p>

              </div>

            </div>

          </div>
        )}

        {/* PRESENTATION */}

        {step === 4 &&
          presentationGenerated && (
            <div>

              <div className="presentation-header">

                <div>

                  <div className="eyebrow">
                    PRESENTATION BUILDER
                  </div>

                  <h1>
                    {project.name ||
                      "Research Presentation"}
                  </h1>

                  <p>
                    {slides.length} slides •{" "}
                    {presentationStyle} style
                  </p>

                </div>

                <button
                  className="secondary-button"
                  onClick={() =>
                    setPresentationGenerated(
                      false
                    )
                  }
                >
                  ← Edit Settings
                </button>

              </div>

              <div className="presentation-workspace">

                <div className="slide-sidebar">

                  <div className="sidebar-title">
                    SLIDES
                  </div>

                  {slides.map(
                    (slide, index) => (
                      <button
                        className={
                          index === currentSlide
                            ? "slide-thumbnail active"
                            : "slide-thumbnail"
                        }
                        key={index}
                        onClick={() =>
                          setCurrentSlide(index)
                        }
                      >

                        <span>
                          {String(
                            index + 1
                          ).padStart(2, "0")}
                        </span>

                        <div>
                          {slide.title}
                        </div>

                      </button>
                    )
                  )}

                </div>

                <div className="slide-main">

                  <div className="slide-toolbar">

                    <div>
                      Slide{" "}
                      {currentSlide + 1} of{" "}
                      {slides.length}
                    </div>

                    <div className="slide-controls">

                      <button
                        disabled={
                          currentSlide === 0
                        }
                        onClick={() =>
                          setCurrentSlide(
                            currentSlide - 1
                          )
                        }
                      >
                        ←
                      </button>

                      <button
                        disabled={
                          currentSlide ===
                          slides.length - 1
                        }
                        onClick={() =>
                          setCurrentSlide(
                            currentSlide + 1
                          )
                        }
                      >
                        →
                      </button>

                    </div>

                  </div>

                  <div className="ppt-canvas">

                    {renderPresentationSlide(
                      slides[currentSlide]
                    )}

                  </div>

                  <div className="edit-slide">

                    <div className="eyebrow">
                      EDIT SLIDE
                    </div>

                    <label>
                      Slide title
                    </label>

                    <input
                      value={
                        slides[currentSlide]
                          ?.title || ""
                      }
                      onChange={(event) =>
                        updateSlide(
                          "title",
                          event.target.value
                        )
                      }
                    />

                    {slides[currentSlide]
                      ?.subtitle !==
                      undefined && (
                      <>
                        <label>
                          Subtitle
                        </label>

                        <input
                          value={
                            slides[currentSlide]
                              ?.subtitle || ""
                          }
                          onChange={(event) =>
                            updateSlide(
                              "subtitle",
                              event.target.value
                            )
                          }
                        />
                      </>
                    )}

                  </div>

                </div>

              </div>

              <div className="export-area">

                <div>
                  <div className="eyebrow">
                    FINAL STEP
                  </div>

                  <h2>
                    Your research story is ready.
                  </h2>

                  <p>
                    Export functionality will be
                    connected next. The prototype
                    currently lets you build,
                    review and edit the entire deck.
                  </p>
                </div>

                <button
                  className="export-button"
                  disabled
                >
                  Export to PowerPoint — Coming Next
                </button>

              </div>

            </div>
          )}

      </div>

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
          background: white;
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
          max-width: 1150px;
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
          background: white;
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
        textarea,
        select {
          width: 100%;
          padding: 14px;
          border: 1px solid #d1d5db;
          border-radius: 10px;
          font-size: 16px;
          font-family: Arial, sans-serif;
          background: white;
        }

        textarea {
          resize: vertical;
        }

        .primary-button,
        .dark-button,
        .presentation-button {
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
          background: white;
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
          background: white;
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

        .presentation-button {
          background: white;
          color: #111827;
          width: auto;
          padding: 14px 22px;
        }

        .builder-layout {
          display: grid;
          grid-template-columns: 1.2fr 0.8fr;
          gap: 25px;
        }

        .builder-description {
          color: #6b7280;
          margin-bottom: 30px;
        }

        .builder-preview {
          background: #111827;
          border-radius: 18px;
          padding: 30px;
          color: white;
        }

        .preview-label {
          color: #a5b4fc;
          font-size: 12px;
          font-weight: 800;
          letter-spacing: 1px;
          margin-bottom: 20px;
        }

        .mini-slide {
          background: white;
          color: #111827;
          aspect-ratio: 16 / 9;
          border-radius: 8px;
          padding: 25px;
          display: flex;
          flex-direction: column;
          justify-content: center;
        }

        .mini-brand {
          color: #4f46e5;
          font-size: 9px;
          font-weight: 800;
          letter-spacing: 1px;
        }

        .mini-title {
          font-size: 25px;
          font-weight: 800;
          margin-top: 15px;
        }

        .mini-subtitle {
          color: #6b7280;
          font-size: 12px;
          margin-top: 8px;
        }

        .mini-lines {
          margin-top: 25px;
        }

        .mini-lines span {
          display: block;
          height: 4px;
          background: #e5e7eb;
          margin-bottom: 7px;
          border-radius: 4px;
        }

        .builder-preview > p {
          color: #9ca3af;
          line-height: 1.5;
        }

        .presentation-header {
          display: flex;
          justify-content: space-between;
          align-items: flex-end;
          margin-bottom: 30px;
        }

        .presentation-header h1 {
          margin-bottom: 5px;
        }

        .presentation-header p {
          color: #6b7280;
        }

        .secondary-button {
          background: white;
          border: 1px solid #d1d5db;
          padding: 12px 18px;
          border-radius: 9px;
          cursor: pointer;
          font-weight: 700;
        }

        .presentation-workspace {
          display: grid;
          grid-template-columns: 260px 1fr;
          gap: 20px;
        }

        .slide-sidebar {
          background: white;
          border: 1px solid #e5e7eb;
          border-radius: 15px;
          padding: 15px;
          height: fit-content;
        }

        .sidebar-title {
          color: #6b7280;
          font-size: 11px;
          font-weight: 800;
          letter-spacing: 1px;
          padding: 5px 8px 15px;
        }

        .slide-thumbnail {
          width: 100%;
          text-align: left;
          display: flex;
          gap: 10px;
          align-items: flex-start;
          padding: 12px 10px;
          margin-bottom: 5px;
          border: 1px solid transparent;
          border-radius: 9px;
          background: transparent;
          cursor: pointer;
          color: #4b5563;
        }

        .slide-thumbnail span {
          color: #9ca3af;
          font-size: 11px;
          font-weight: 800;
        }

        .slide-thumbnail div {
          font-size: 12px;
          line-height: 1.4;
        }

        .slide-thumbnail.active {
          background: #eef2ff;
          border-color: #c7d2fe;
          color: #3730a3;
        }

        .slide-main {
          min-width: 0;
        }

        .slide-toolbar {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 12px;
          color: #6b7280;
          font-size: 13px;
        }

        .slide-controls {
          display: flex;
          gap: 6px;
        }

        .slide-controls button {
          width: 34px;
          height: 30px;
          border: 1px solid #d1d5db;
          background: white;
          border-radius: 6px;
          cursor: pointer;
        }

        .slide-controls button:disabled {
          opacity: 0.4;
          cursor: default;
        }

        .ppt-canvas {
          background: #d1d5db;
          padding: 25px;
          border-radius: 15px;
        }

        .ppt-title-slide,
        .ppt-content-slide {
          background: white;
          aspect-ratio: 16 / 9;
          border-radius: 4px;
          padding: 55px;
          box-shadow:
            0 10px 30px rgba(0,0,0,0.08);
        }

        .ppt-title-slide {
          display: flex;
          flex-direction: column;
          justify-content: center;
          position: relative;
        }

        .ppt-brand,
        .ppt-small-label {
          color: #4f46e5;
          font-size: 11px;
          font-weight: 800;
          letter-spacing: 1.5px;
        }

        .ppt-title-slide h1 {
          font-size: 44px;
          max-width: 800px;
          margin-top: 25px;
          margin-bottom: 15px;
        }

        .ppt-title-slide p {
          color: #6b7280;
          font-size: 19px;
        }

        .ppt-footer {
          position: absolute;
          bottom: 35px;
          left: 55px;
          color: #9ca3af;
          font-size: 11px;
        }

        .ppt-content-slide h1 {
          font-size: 32px;
          margin: 12px 0 8px;
        }

        .ppt-subtitle {
          color: #6b7280;
          font-size: 14px;
          margin-top: 0;
        }

        .ppt-bullets {
          margin-top: 35px;
        }

        .ppt-bullet {
          display: flex;
          gap: 18px;
          padding: 14px 0;
          border-bottom: 1px solid #e5e7eb;
          font-size: 15px;
          line-height: 1.5;
        }

        .ppt-bullet-number {
          color: #4f46e5;
          font-weight: 800;
          min-width: 25px;
        }

        .ppt-bullet-dot {
          color: #4f46e5;
          min-width: 20px;
        }

        .ppt-theme-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 15px;
          margin-top: 30px;
        }

        .ppt-theme {
          border: 1px solid #e5e7eb;
          border-radius: 10px;
          padding: 18px;
        }

        .ppt-theme-number {
          color: #4f46e5;
          font-size: 20px;
          font-weight: 800;
        }

        .ppt-theme h3 {
          margin: 8px 0;
          font-size: 16px;
        }

        .ppt-theme span {
          font-size: 9px;
          color: #16a34a;
          font-weight: 800;
        }

        .ppt-segment-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 15px;
          margin-top: 30px;
        }

        .ppt-segment {
          background: #f9fafb;
          border-radius: 10px;
          padding: 18px;
        }

        .ppt-segment h3 {
          margin-top: 0;
          font-size: 15px;
        }

        .ppt-segment p {
          color: #6b7280;
          font-size: 12px;
          line-height: 1.5;
        }

        .ppt-quote-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 15px;
          margin-top: 30px;
        }

        .ppt-quote {
          background: #f9fafb;
          padding: 18px;
          border-radius: 10px;
          position: relative;
        }

        .quote-mark {
          color: #4f46e5;
          font-size: 30px;
          font-weight: 800;
        }

        .ppt-quote p {
          font-size: 13px;
          line-height: 1.5;
          font-style: italic;
        }

        .ppt-quote small {
          color: #6b7280;
          font-size: 9px;
        }

        .edit-slide {
          background: white;
          border: 1px solid #e5e7eb;
          border-radius: 12px;
          padding: 20px;
          margin-top: 15px;
        }

        .edit-slide label {
          margin-top: 15px;
          font-size: 13px;
        }

        .edit-slide input {
          font-size: 14px;
        }

        .export-area {
          margin-top: 30px;
          background: #111827;
          color: white;
          border-radius: 18px;
          padding: 30px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 30px;
        }

        .export-area h2 {
          margin-bottom: 8px;
        }

        .export-area p {
          color: #9ca3af;
          max-width: 650px;
          line-height: 1.5;
        }

        .export-button {
          background: #374151;
          color: #9ca3af;
          border: none;
          padding: 14px 20px;
          border-radius: 9px;
          font-weight: 700;
          white-space: nowrap;
        }

        @media (max-width: 850px) {

          .builder-layout {
            grid-template-columns: 1fr;
          }

          .presentation-workspace {
            grid-template-columns: 1fr;
          }

          .slide-sidebar {
            display: flex;
            overflow-x: auto;
          }

          .slide-thumbnail {
            min-width: 150px;
          }

          .export-area {
            flex-direction: column;
            align-items: flex-start;
          }

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

          .ppt-title-slide,
          .ppt-content-slide {
            padding: 25px;
          }

          .ppt-title-slide h1 {
            font-size: 28px;
          }

          .ppt-content-slide h1 {
            font-size: 24px;
          }

          .ppt-canvas {
            padding: 10px;
          }

        }

      `}</style>

    </main>
  );
}
