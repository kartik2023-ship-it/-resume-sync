import { NextRequest, NextResponse } from "next/server";
import { fillATSTemplate } from "@/app/lib/fillATSTemplate";
import type { StructuredResume } from "@/app/lib/types";

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const structured = body.structured as StructuredResume;

    if (!structured) {
      return NextResponse.json({ error: "Missing structured resume data" }, { status: 400 });
    }

    const buffer = fillATSTemplate(structured);

    return new NextResponse(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": 'attachment; filename="resume-ats.docx"',
      },
    });
  } catch (err) {
    console.error("ATS download error:", err);
    return NextResponse.json({ error: "Failed to generate ATS document" }, { status: 500 });
  }
}
