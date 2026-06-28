import {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  Packer,
  Paragraph,
  TabStopType,
  TextRun,
} from "docx";
import type { StructuredResume } from "./types";

const FONT = "Calibri";
const MARGIN = convertInchesToTwip(0.75);
// Content width: 8.5 - 0.75 - 0.75 = 7 inches
const CENTER_TAB = convertInchesToTwip(3.5);
const RIGHT_TAB = convertInchesToTwip(7);

const METRIC_RE =
  /(\$[\d,]+(?:\.\d+)?[KkMmBb]?|Rs\.?\s*[\d,]+(?:\.\d+)?[KkMmBb]?|₹[\d,]+(?:\.\d+)?[KkMmBb]?|\d+(?:,\d+)*(?:\.\d+)?%|\b\d+(?:,\d+)*(?:\.\d+)?[KkMmBb]?\b)/g;

function boldMetrics(text: string, size = 20): TextRun[] {
  const runs: TextRun[] = [];
  let last = 0;
  let m: RegExpExecArray | null;
  METRIC_RE.lastIndex = 0;

  while ((m = METRIC_RE.exec(text)) !== null) {
    if (m.index > last) {
      runs.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size }));
    }
    runs.push(new TextRun({ text: m[0], bold: true, font: FONT, size }));
    last = m.index + m[0].length;
  }

  if (last < text.length) {
    runs.push(new TextRun({ text: text.slice(last), font: FONT, size }));
  }

  return runs.length ? runs : [new TextRun({ text, font: FONT, size })];
}

function sectionHeader(title: string): Paragraph {
  return new Paragraph({
    border: {
      bottom: { color: "000000", style: BorderStyle.SINGLE, size: 6, space: 1 },
    },
    spacing: { before: 160, after: 80 },
    children: [
      new TextRun({ text: title.toUpperCase(), bold: true, size: 22, font: FONT }),
    ],
  });
}

export async function generateATSDocx(data: StructuredResume): Promise<Blob> {
  const children: Paragraph[] = [];

  // Name
  children.push(
    new Paragraph({
      children: [new TextRun({ text: data.name, bold: true, size: 32, font: FONT })],
      spacing: { after: 40 },
    })
  );

  // Contact line
  const contactParts = [
    data.contact.phone,
    data.contact.email,
    data.contact.linkedin,
    data.contact.website,
  ].filter(Boolean);

  if (contactParts.length) {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: contactParts.join(" | "), size: 20, font: FONT })],
        spacing: { after: 60 },
      })
    );
  }

  // Horizontal rule after header
  children.push(
    new Paragraph({
      border: {
        bottom: { color: "000000", style: BorderStyle.SINGLE, size: 6, space: 1 },
      },
      children: [],
      spacing: { after: 80 },
    })
  );

  // Summary
  if (data.summary) {
    children.push(sectionHeader("SUMMARY"));
    children.push(
      new Paragraph({
        children: [new TextRun({ text: data.summary, size: 20, font: FONT })],
        spacing: { after: 80 },
      })
    );
  }

  // Experience
  if (data.experience.length) {
    children.push(sectionHeader("EXPERIENCE"));
    for (const exp of data.experience) {
      children.push(
        new Paragraph({
          tabStops: [
            { type: TabStopType.CENTER, position: CENTER_TAB },
            { type: TabStopType.RIGHT, position: RIGHT_TAB },
          ],
          children: [
            new TextRun({ text: exp.company, bold: true, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: exp.role, bold: true, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: exp.dates, size: 20, font: FONT }),
          ],
          spacing: { before: 100, after: 40 },
        })
      );
      for (const cat of exp.categories) {
        for (const bullet of cat.bullets) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({ text: "• ", size: 20, font: FONT }),
                ...boldMetrics(bullet),
              ],
              indent: { left: convertInchesToTwip(0.2) },
              spacing: { after: 40 },
            })
          );
        }
      }
    }
  }

  // Projects
  if (data.projects.length) {
    children.push(sectionHeader("PROJECTS"));
    for (const proj of data.projects) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({ text: proj.name, bold: true, size: 20, font: FONT }),
            new TextRun({ text: " — " + proj.description, size: 20, font: FONT }),
          ],
          spacing: { after: 60 },
        })
      );
    }
  }

  // Internships
  if (data.internships.length) {
    children.push(sectionHeader("INTERNSHIPS"));
    for (const intern of data.internships) {
      children.push(
        new Paragraph({
          tabStops: [
            { type: TabStopType.CENTER, position: CENTER_TAB },
            { type: TabStopType.RIGHT, position: RIGHT_TAB },
          ],
          children: [
            new TextRun({ text: intern.company, bold: true, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: intern.role, bold: true, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: intern.dates, size: 20, font: FONT }),
          ],
          spacing: { before: 100, after: 40 },
        })
      );
      for (const cat of intern.categories) {
        for (const bullet of cat.bullets) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({ text: "• ", size: 20, font: FONT }),
                ...boldMetrics(bullet),
              ],
              indent: { left: convertInchesToTwip(0.2) },
              spacing: { after: 40 },
            })
          );
        }
      }
    }
  }

  // Education
  if (data.education.length) {
    children.push(sectionHeader("EDUCATION"));
    for (const edu of data.education) {
      children.push(
        new Paragraph({
          tabStops: [
            { type: TabStopType.CENTER, position: CENTER_TAB },
            { type: TabStopType.RIGHT, position: RIGHT_TAB },
          ],
          children: [
            new TextRun({ text: edu.degree, bold: true, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: edu.institute, size: 20, font: FONT }),
            new TextRun({ text: "\t" }),
            new TextRun({ text: edu.year, size: 20, font: FONT }),
          ],
          spacing: { after: 60 },
        })
      );
    }
  }

  // Achievements
  if (data.achievements.length) {
    children.push(sectionHeader("ACHIEVEMENTS"));
    for (const ach of data.achievements) {
      const text = ach.year ? `${ach.achievement} (${ach.year})` : ach.achievement;
      children.push(
        new Paragraph({
          children: [
            new TextRun({ text: "• ", size: 20, font: FONT }),
            ...boldMetrics(text),
          ],
          indent: { left: convertInchesToTwip(0.2) },
          spacing: { after: 40 },
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        children,
      },
    ],
  });

  return Packer.toBlob(doc);
}
