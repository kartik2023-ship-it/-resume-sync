import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import mammoth from "mammoth";

const SYSTEM_PROMPT = `You are a professional resume writer and ATS expert.

Rewrite the provided resume to better match the job description. Rules:
- NEVER add any skill, tool, achievement, or experience not present in the original resume
- Only rephrase, reorder, and align existing content with JD keywords and language
- Preserve all factual details exactly (dates, company names, job titles, metrics)

Format the rewritten resume clearly using plain text:
- Section headers in ALL CAPS (e.g. SUMMARY, EXPERIENCE, EDUCATION, SKILLS)
- A line of hyphens under each header (e.g. ----------)
- Job/education entries on their own line, e.g. "Job Title | Company | Start – End"
- Bullet points using "• " for responsibilities and achievements
- One blank line between sections

After rewriting, score how well the result matches the JD.

Return ONLY a valid JSON object — no markdown fences, no explanation, nothing else:
{
  "resume": "<full formatted resume with \\n for newlines>",
  "ats": {
    "score": <integer 0-100, overall ATS match>,
    "keywordMatch": <integer 0-100, % of important JD keywords present>,
    "skillsCoverage": <integer 0-100, % of required/preferred skills covered>,
    "suggestions": ["<specific actionable suggestion>", "<specific actionable suggestion>", "<specific actionable suggestion>"]
  }
}`;

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("resume") as File | null;
    const jd = formData.get("jd") as string | null;

    if (!file || !jd) {
      return NextResponse.json(
        { error: "Missing resume file or job description" },
        { status: 400 }
      );
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);

    let resumeText = "";

    if (
      file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
      file.name.endsWith(".docx")
    ) {
      const result = await mammoth.extractRawText({ buffer });
      resumeText = result.value;
    } else if (file.type === "application/pdf" || file.name.endsWith(".pdf")) {
      const { PDFParse } = await import("pdf-parse");
      const parser = new PDFParse({ data: buffer });
      const data = await parser.getText();
      resumeText = data.text;
    } else {
      return NextResponse.json(
        { error: "Unsupported file type. Please upload a PDF or DOCX." },
        { status: 400 }
      );
    }

    if (!resumeText.trim()) {
      return NextResponse.json(
        { error: "Could not extract text from the resume." },
        { status: 400 }
      );
    }

    const client = new Anthropic({
      apiKey: process.env.ANTHROPIC_API_KEY,
    });

    const response = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 4096,
      system: SYSTEM_PROMPT,
      messages: [
        {
          role: "user",
          content: `ORIGINAL RESUME:\n\n${resumeText}\n\n---\n\nJOB DESCRIPTION:\n\n${jd}`,
        },
      ],
    });

    const textBlock = response.content.find((b) => b.type === "text");
    if (!textBlock || textBlock.type !== "text") {
      return NextResponse.json(
        { error: "No text response from AI" },
        { status: 500 }
      );
    }

    // Strip markdown code fences if Claude wraps the JSON
    const raw = textBlock.text
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/, "")
      .trim();

    try {
      const parsed = JSON.parse(raw) as {
        resume: string;
        ats: {
          score: number;
          keywordMatch: number;
          skillsCoverage: number;
          suggestions: string[];
        };
      };
      return NextResponse.json({ resume: parsed.resume, ats: parsed.ats });
    } catch {
      return NextResponse.json({ resume: raw, ats: null });
    }
  } catch (err) {
    console.error("Rewrite error:", err);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
