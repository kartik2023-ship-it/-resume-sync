import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import mammoth from "mammoth";
import type { StructuredResume } from "@/app/lib/types";

const SYSTEM_PROMPT = `You are a professional resume writer and ATS expert.

Rewrite the provided resume to better match the job description. Rules:
- NEVER add any skill, tool, achievement, or experience not present in the original resume
- Only rephrase, reorder, and align existing content with JD keywords and language
- Preserve all factual details exactly (dates, company names, job titles, metrics)
- If a section (internships, projects, achievements) does not exist in the original resume, return an empty array for it
- Group bullets under functional categories (e.g. Engineering, Product, Leadership) — use a single category if the work was uniform

Return ONLY a valid JSON object — no markdown fences, no explanation, nothing else:
{
  "name": "<candidate full name>",
  "contact": {
    "phone": "<phone number or empty string>",
    "email": "<email address or empty string>",
    "linkedin": "<LinkedIn URL or handle or empty string>",
    "website": "<website or portfolio URL or empty string>"
  },
  "summary": "<2-3 sentence professional summary tailored to the JD>",
  "experience": [
    {
      "company": "<company name>",
      "role": "<job title>",
      "dates": "<start – end dates>",
      "categories": [
        {
          "label": "<functional category, e.g. Engineering, Product, Leadership>",
          "bullets": ["<rewritten achievement or responsibility>"]
        }
      ]
    }
  ],
  "projects": [
    {
      "name": "<project name>",
      "description": "<one-line description>"
    }
  ],
  "internships": [
    {
      "company": "<company name>",
      "role": "<internship title>",
      "dates": "<start – end dates>",
      "categories": [
        {
          "label": "<functional category>",
          "bullets": ["<rewritten achievement or responsibility>"]
        }
      ]
    }
  ],
  "education": [
    {
      "degree": "<degree name and major>",
      "institute": "<institution name>",
      "year": "<graduation year or expected year>"
    }
  ],
  "achievements": [
    {
      "achievement": "<achievement description>",
      "year": "<year or empty string>"
    }
  ],
  "ats": {
    "score": <integer 0-100, overall ATS match>,
    "keywordMatch": <integer 0-100, % of important JD keywords present>,
    "skillsCoverage": <integer 0-100, % of required/preferred skills covered>,
    "suggestions": ["<specific actionable suggestion>", "<specific actionable suggestion>", "<specific actionable suggestion>"]
  }
}`;

function structuredToText(data: StructuredResume): string {
  const lines: string[] = [];

  lines.push(data.name);

  const contactParts = [
    data.contact.phone,
    data.contact.email,
    data.contact.linkedin,
    data.contact.website,
  ].filter(Boolean);
  if (contactParts.length) lines.push(contactParts.join(" | "));

  lines.push("");

  if (data.summary) {
    lines.push("SUMMARY");
    lines.push("----------");
    lines.push(data.summary);
    lines.push("");
  }

  if (data.experience.length) {
    lines.push("EXPERIENCE");
    lines.push("----------");
    for (const exp of data.experience) {
      lines.push(`${exp.role} | ${exp.company} | ${exp.dates}`);
      for (const cat of exp.categories) {
        for (const bullet of cat.bullets) lines.push(`• ${bullet}`);
      }
      lines.push("");
    }
  }

  if (data.projects.length) {
    lines.push("PROJECTS");
    lines.push("----------");
    for (const proj of data.projects) {
      lines.push(`${proj.name} — ${proj.description}`);
    }
    lines.push("");
  }

  if (data.internships.length) {
    lines.push("INTERNSHIPS");
    lines.push("----------");
    for (const intern of data.internships) {
      lines.push(`${intern.role} | ${intern.company} | ${intern.dates}`);
      for (const cat of intern.categories) {
        for (const bullet of cat.bullets) lines.push(`• ${bullet}`);
      }
      lines.push("");
    }
  }

  if (data.education.length) {
    lines.push("EDUCATION");
    lines.push("----------");
    for (const edu of data.education) {
      lines.push(`${edu.degree} | ${edu.institute} | ${edu.year}`);
    }
    lines.push("");
  }

  if (data.achievements.length) {
    lines.push("ACHIEVEMENTS");
    lines.push("----------");
    for (const ach of data.achievements) {
      const text = ach.year ? `${ach.achievement} (${ach.year})` : ach.achievement;
      lines.push(`• ${text}`);
    }
    lines.push("");
  }

  return lines.join("\n").trim();
}

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
      // eslint-disable-next-line @typescript-eslint/no-require-imports
      const pdfParse = require("pdf-parse");
      const data = await pdfParse(buffer);
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
      max_tokens: 8192,
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
      return NextResponse.json({ error: "No text response from AI" }, { status: 500 });
    }

    const raw = textBlock.text
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/, "")
      .trim();

    try {
      const parsed = JSON.parse(raw) as StructuredResume & {
        ats: { score: number; keywordMatch: number; skillsCoverage: number; suggestions: string[] };
      };

      const structured: StructuredResume = {
        name: parsed.name ?? "",
        contact: parsed.contact ?? { phone: "", email: "", linkedin: "", website: "" },
        summary: parsed.summary ?? "",
        experience: parsed.experience ?? [],
        projects: parsed.projects ?? [],
        internships: parsed.internships ?? [],
        education: parsed.education ?? [],
        achievements: parsed.achievements ?? [],
      };

      const resume = structuredToText(structured);

      return NextResponse.json({ resume, structured, ats: parsed.ats });
    } catch {
      return NextResponse.json({ resume: raw, structured: null, ats: null });
    }
  } catch (err) {
    console.error("Rewrite error:", err);
    return NextResponse.json({ error: "Internal server error" }, { status: 500 });
  }
}
