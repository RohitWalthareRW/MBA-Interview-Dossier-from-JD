# MBA-Interview-Dossier-from-JD
An end-to-end multimodal AI research agent that transforms any Job Description (JD) and optional Candidate Résumé/CV-into a structured, source-backed 10-Part MBA Interview Preparation Dossier in under two minutes. 
Built with Python, Streamlit, and the Google Gemini API, the agent automates deep company intelligence, financial analysis, industry frameworks, multi-window news tracking, KPI mapping, and tailored STAR interview preparation while enforcing strict separation between verified facts and analytical inferences.

Key Features

Multimodal Document Ingestion: Accepts JDs in .pdf, .pptx, .docx, .png, .jpg, .txt, or pasted raw text. Routes PDFs and images natively through the Gemini File API and extracts slide text and tables cleanly from PowerPoint presentations.

Live Web Research & Grounded Citations: Uses Google Search grounding anchored to the dynamic runtime date to gather current company financials, leadership details, competitor benchmarks, and regulatory developments—extracting verified URLs directly from grounding metadata to eliminate fabricated links.

Multi-Window News Intelligence: Tracks and analyzes material company and industry developments across four time horizons (Last 30 days, 3–6 months, 6–12 months, and 1–2 years) and links each headline directly to role relevance and interview talking points.

Epistemic Rigor & Hallucination Guardrails: Enforces structured intermediate schemas via Pydantic, tagging claims clearly across [Fact], [Management Statement], [Estimate], [Inference], and [Recommendation], and flagging entity ambiguity or undisclosed private-company financials honestly.

Conditional Candidate Fit & STAR Synthesis:

With CV Uploaded: Maps documented candidate achievements and metrics directly to JD requirements and drafts tailored STAR stories with probing follow-up questions.

Without CV Uploaded: Generates a structured Candidate Evidence Worksheet with placeholders without inventing personal experience.

Prompt-Injection Sandboxing: Wraps external document text in explicit untrusted-data boundaries to prevent indirect prompt injection.

Multi-Model Tier Selector & One-Click Exports: Switch seamlessly between Gemini Flash, Gemini Pro, and Gemini Argon tiers in the UI, and export the final dossier as a formatted Word Document (.docx), Markdown (.md), or Structured JSON (.json).

10-Part Dossier Output Structure

One-Page Final Revision Sheet: 30–45 second company pitch, business model summary, top competitors, 3 key news items, 3 role KPIs, 2 strategic recommendations, and top 5 priority questions for rapid last-minute review.

Executive Briefing: High-level synthesis and top 10–15 memorable talking points.

Company Dossier: Founding history, ownership, leadership, monetization & unit economics, multi-year financial snapshot table, culture signals, and governance/risk profile.

Industry & Competitive Landscape: Market size/CAGR, demand drivers, competitor comparison, Porter's Five Forces, PESTLE, Value Chain analysis, and macro/regulatory transmission mechanisms.

Recent News Tracker: Structured table capturing publication/event dates, summaries, verified links, business implications, risks, and JD relevance.

JD & Role Analysis: Responsibility-to-company-priority mapping table, deep-dive role KPIs, and a proposed first-90-days execution plan.

Strategic Assessment & Recommendations: Evidence-backed growth and efficiency/risk recommendations with implementation steps, trade-offs, and monitoring KPIs.

Candidate Fit & Personal Stories: CV-backed fit matrix and STAR stories (or structured preparation worksheet) plus core behavioral frameworks.

Interview Question Bank: Categorized Company, Industry, Role/Case Judgment, and HR questions with answer frameworks and common pitfalls.

Questions for the Interviewer & Verified Bibliography: 4–6 research-informed questions to ask the panel, followed by a consolidated source register and research limitations log.

Tech Stack

LLM & Grounding: Google Gemini API (google-genai), Google Search Grounding, Gemini File API

Frontend UI: Streamlit

Data Validation: Pydantic

Document Parsing & Export: python-pptx, python-docx, Pillow

Quickstart & Local Setup

1. Prerequisites

Python 3.10 or higher

A Google Gemini API key from Google AI Studio

2. Clone the Repository

git clone https://github.com/yourusername/mba-interview-research-agent.git
cd mba-interview-research-agent


3. Install Dependencies

pip install -r requirements.txt


4. Run the Application

python -m streamlit run app.py


The web interface will open automatically in your browser at http://localhost:8501. Enter your Gemini API key in the sidebar, upload a JD (and optional CV), and click Generate Interview Dossier.

Project Structure

mba-interview-research-agent/
├── app.py              # Complete Streamlit UI, multimodal ingestion, grounded research & export pipeline
├── requirements.txt    # Python package dependencies
└── README.md           # Project documentation


Privacy & Security Notes

No Hardcoded Secrets: API keys are entered dynamically at runtime via the Streamlit sidebar or read from local environment variables (GEMINI_API_KEY).

Ephemeral File Cleanup: Uploaded PDFs and images sent to the Gemini File API are automatically deleted from remote servers and local temporary storage immediately after each run completes.
