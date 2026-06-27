import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import mammoth from "mammoth";

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
      // Dynamically import pdf-parse to avoid Next.js build issues
      const pdfParse = (await import("pdf-parse")).default;
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
      max_tokens: 4096,
      system:
        "Rewrite this resume to match the JD keywords and language. Never add any skill, tool, achievement or experience not present in the original resume. Only rephrase and reorder what exists.",
      messages: [
        {
          role: "user",
          content: `Here is the original resume:\n\n${resumeText}\n\n---\n\nHere is the job description:\n\n${jd}\n\n---\n\nPlease rewrite the resume to better match the job description, using its keywords and language style. Do not invent or add anything that is not already in the resume.`,
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

    return NextResponse.json({ result: textBlock.text });
  } catch (err) {
    console.error("Rewrite error:", err);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
