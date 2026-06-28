import { NextRequest, NextResponse } from "next/server";
import { fillVisualTemplate } from "@/app/lib/fillVisualTemplate";
import type { StructuredResume } from "@/app/lib/types";

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const structured = body.structured as StructuredResume;

    if (!structured) {
      return NextResponse.json({ error: "Missing structured resume data" }, { status: 400 });
    }

    const buffer = fillVisualTemplate(structured);

    return new NextResponse(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": 'attachment; filename="resume-visual.docx"',
      },
    });
  } catch (err) {
    console.error("Visual download error:", err);
    return NextResponse.json({ error: "Failed to generate visual document" }, { status: 500 });
  }
}
