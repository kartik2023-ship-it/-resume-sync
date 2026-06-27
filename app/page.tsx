"use client";

import { useState } from "react";
import { Document, Packer, Paragraph, TextRun } from "docx";

interface ATSResult {
  score: number;
  keywordMatch: number;
  skillsCoverage: number;
  suggestions: string[];
}

function scoreTheme(score: number) {
  if (score >= 80)
    return {
      ring: "ring-green-200",
      bg: "bg-green-50",
      text: "text-green-700",
      bar: "bg-green-500",
      badge: "bg-green-100 text-green-800",
    };
  if (score >= 60)
    return {
      ring: "ring-amber-200",
      bg: "bg-amber-50",
      text: "text-amber-700",
      bar: "bg-amber-500",
      badge: "bg-amber-100 text-amber-800",
    };
  return {
    ring: "ring-red-200",
    bg: "bg-red-50",
    text: "text-red-700",
    bar: "bg-red-500",
    badge: "bg-red-100 text-red-800",
  };
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [jd, setJd] = useState("");
  const [resume, setResume] = useState("");
  const [ats, setAts] = useState<ATSResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!file || !jd.trim()) {
      setError("Please upload a resume and paste the job description.");
      return;
    }
    setError("");
    setLoading(true);
    setResume("");
    setAts(null);

    try {
      const formData = new FormData();
      formData.append("resume", file);
      formData.append("jd", jd);

      const res = await fetch("/api/rewrite", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error ?? "Something went wrong.");
      } else {
        setResume(data.resume ?? "");
        setAts(data.ats ?? null);
      }
    } catch {
      setError("Network error. Please try again.");
    } finally {
      setLoading(false);
    }
  }

  async function handleDownload() {
    if (!resume) return;

    const paragraphs = resume.split("\n").map(
      (line) => new Paragraph({ children: [new TextRun(line)] })
    );

    const doc = new Document({ sections: [{ children: paragraphs }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "rewritten-resume.docx";
    a.click();
    URL.revokeObjectURL(url);
  }

  const theme = ats ? scoreTheme(ats.score) : null;
  const hasOutput = !!(resume || loading);

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4">
        <div className="max-w-7xl mx-auto flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center flex-shrink-0">
            <span className="text-white text-sm font-bold">R</span>
          </div>
          <div>
            <h1 className="text-base font-semibold text-gray-900 leading-tight">
              ResumeSync
            </h1>
            <p className="text-xs text-gray-400">AI-powered resume tailoring</p>
          </div>
        </div>
      </header>

      {/* Two-column layout */}
      <main className="max-w-7xl mx-auto px-4 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start">
          {/* ── Left column: inputs ── */}
          <form onSubmit={handleSubmit} className="space-y-4">
            {/* Upload card */}
            <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
              <p className="text-xs font-semibold text-gray-500 uppercase tracking-widest mb-3">
                Resume
              </p>
              <label className="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-gray-200 rounded-lg cursor-pointer hover:border-indigo-300 hover:bg-indigo-50/50 transition-colors">
                {file ? (
                  <div className="flex items-center gap-2 text-indigo-600">
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                    </svg>
                    <span className="text-sm font-medium truncate max-w-xs">{file.name}</span>
                  </div>
                ) : (
                  <div className="flex flex-col items-center gap-1 text-gray-400">
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                    </svg>
                    <span className="text-sm">Upload PDF or DOCX</span>
                    <span className="text-xs">Click to browse</span>
                  </div>
                )}
                <input
                  type="file"
                  accept=".pdf,.docx"
                  className="hidden"
                  onChange={(e) => setFile(e.target.files?.[0] ?? null)}
                />
              </label>
            </div>

            {/* JD card */}
            <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
              <p className="text-xs font-semibold text-gray-500 uppercase tracking-widest mb-3">
                Job Description
              </p>
              <textarea
                value={jd}
                onChange={(e) => setJd(e.target.value)}
                rows={14}
                placeholder="Paste the job description here…"
                className="w-full text-sm text-gray-800 placeholder-gray-400 resize-none focus:outline-none leading-relaxed"
              />
            </div>

            {error && (
              <p className="text-sm text-red-600 bg-red-50 border border-red-100 rounded-lg px-4 py-2.5">
                {error}
              </p>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-xl bg-indigo-600 px-4 py-3 text-sm font-semibold text-white hover:bg-indigo-500 active:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            >
              {loading ? "Processing…" : "Rewrite Resume"}
            </button>
          </form>

          {/* ── Right column: output ── */}
          <div className="space-y-4">
            {/* Loading state */}
            {loading && (
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-10 flex flex-col items-center justify-center gap-4">
                <div className="w-10 h-10 rounded-full border-4 border-indigo-100 border-t-indigo-600 animate-spin" />
                <div className="text-center">
                  <p className="text-sm font-medium text-gray-700">
                    Analyzing and rewriting your resume…
                  </p>
                  <p className="text-xs text-gray-400 mt-1">
                    This usually takes 15–30 seconds
                  </p>
                </div>
              </div>
            )}

            {/* Empty state */}
            {!loading && !hasOutput && (
              <div className="bg-white rounded-xl border border-dashed border-gray-200 p-10 flex flex-col items-center justify-center gap-3 text-center">
                <div className="w-12 h-12 rounded-full bg-gray-50 flex items-center justify-center">
                  <svg className="w-6 h-6 text-gray-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                </div>
                <p className="text-sm text-gray-400 max-w-xs">
                  Your ATS score and rewritten resume will appear here after you submit
                </p>
              </div>
            )}

            {/* ATS Score card */}
            {!loading && ats && theme && (
              <div className={`rounded-xl border shadow-sm p-5 ring-1 ${theme.bg} ${theme.ring}`}>
                <div className="flex items-start justify-between mb-3">
                  <p className="text-xs font-semibold text-gray-500 uppercase tracking-widest">
                    ATS Match Score
                  </p>
                  <span className={`text-4xl font-bold leading-none ${theme.text}`}>
                    {ats.score}
                    <span className="text-base font-normal text-gray-400">/100</span>
                  </span>
                </div>

                {/* Progress bar */}
                <div className="w-full h-2 bg-white/70 rounded-full mb-4 overflow-hidden">
                  <div
                    className={`h-2 rounded-full ${theme.bar} transition-all duration-700`}
                    style={{ width: `${ats.score}%` }}
                  />
                </div>

                {/* Metrics */}
                <div className="grid grid-cols-2 gap-3 mb-4">
                  <div className="bg-white/70 rounded-lg p-3">
                    <p className="text-xs text-gray-500 mb-0.5">Keyword Match</p>
                    <p className={`text-2xl font-bold ${theme.text}`}>
                      {ats.keywordMatch}%
                    </p>
                  </div>
                  <div className="bg-white/70 rounded-lg p-3">
                    <p className="text-xs text-gray-500 mb-0.5">Skills Coverage</p>
                    <p className={`text-2xl font-bold ${theme.text}`}>
                      {ats.skillsCoverage}%
                    </p>
                  </div>
                </div>

                {/* Suggestions */}
                <div className="bg-white/70 rounded-lg p-3">
                  <p className="text-xs font-semibold text-gray-600 uppercase tracking-wide mb-2">
                    Suggestions to improve
                  </p>
                  <ul className="space-y-2">
                    {ats.suggestions.map((s, i) => (
                      <li key={i} className="flex items-start gap-2 text-sm text-gray-700">
                        <span className={`mt-0.5 flex-shrink-0 w-4 h-4 rounded-full flex items-center justify-center text-xs font-semibold ${theme.badge}`}>
                          {i + 1}
                        </span>
                        {s}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            )}

            {/* Rewritten resume card */}
            {!loading && resume && (
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-xs font-semibold text-gray-500 uppercase tracking-widest">
                    Rewritten Resume
                  </p>
                  <button
                    onClick={handleDownload}
                    className="flex items-center gap-1.5 rounded-lg bg-indigo-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-indigo-500 transition-colors"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                    </svg>
                    Download DOCX
                  </button>
                </div>
                <textarea
                  value={resume}
                  onChange={(e) => setResume(e.target.value)}
                  rows={30}
                  className="w-full text-sm text-gray-800 resize-none focus:outline-none font-mono leading-relaxed"
                />
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
