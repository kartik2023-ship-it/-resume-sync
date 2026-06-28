import {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextDirection,
  TextRun,
  VerticalAlign,
  VerticalMergeType,
  WidthType,
} from "docx";
import type { Experience, StructuredResume } from "./types";

const FONT = "Calibri";
const MARGIN = convertInchesToTwip(0.5);

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "auto" };
const NO_BORDERS = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
  insideH: NO_BORDER,
  insideV: NO_BORDER,
};

const LIGHT_GRAY = "E5E7EB";
const DARK_BG = "1F2937";

function sectionHeader(title: string): Paragraph {
  return new Paragraph({
    shading: { type: ShadingType.SOLID, color: DARK_BG, fill: DARK_BG },
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: title.toUpperCase(), bold: true, color: "FFFFFF", size: 22, font: FONT }),
    ],
  });
}

function experienceTable(entries: Experience[]): Table {
  const rows: TableRow[] = [];

  for (const exp of entries) {
    // Company header row spanning both columns
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 2,
            shading: { type: ShadingType.SOLID, color: LIGHT_GRAY, fill: LIGHT_GRAY },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: exp.company, bold: true, size: 22, font: FONT }),
                  new TextRun({ text: "  ", size: 20, font: FONT }),
                  new TextRun({ text: exp.role, size: 20, font: FONT, italics: true }),
                  new TextRun({ text: "  |  " + exp.dates, size: 18, font: FONT, color: "6B7280" }),
                ],
              }),
            ],
          }),
        ],
      })
    );

    // Category rows with rotated label on left
    for (const cat of exp.categories) {
      const bulletCount = cat.bullets.length;
      if (bulletCount === 0) continue;

      for (let i = 0; i < bulletCount; i++) {
        const isFirst = i === 0;
        const cellsInRow: TableCell[] = [];

        if (isFirst) {
          // Category label cell spanning all bullet rows
          const labelCell = new TableCell({
            width: { size: 10, type: WidthType.PERCENTAGE },
            textDirection: TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT,
            verticalAlign: VerticalAlign.CENTER,
            shading: { type: ShadingType.SOLID, color: "F3F4F6", fill: "F3F4F6" },
            // Use verticalMerge RESTART for first, subsequent rows get CONTINUE cells
            verticalMerge: bulletCount > 1 ? VerticalMergeType.RESTART : undefined,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: cat.label.toUpperCase(),
                    bold: true,
                    size: 16,
                    font: FONT,
                    color: "374151",
                  }),
                ],
              }),
            ],
          });
          cellsInRow.push(labelCell);
        } else {
          // Continuation cell for vertical merge
          cellsInRow.push(
            new TableCell({
              width: { size: 10, type: WidthType.PERCENTAGE },
              verticalMerge: VerticalMergeType.CONTINUE,
              children: [new Paragraph("")],
            })
          );
        }

        // Bullet cell
        cellsInRow.push(
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "• " + cat.bullets[i], size: 20, font: FONT }),
                ],
                spacing: { after: 40 },
              }),
            ],
          })
        );

        rows.push(new TableRow({ children: cellsInRow }));
      }
    }

    // Spacer row
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 2,
            children: [new Paragraph({ spacing: { after: 60 } })],
          }),
        ],
      })
    );
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows,
  });
}

export async function generateVisualDocx(data: StructuredResume): Promise<Blob> {
  const currentRole = data.experience[0]?.role ?? data.internships[0]?.role ?? "";
  const currentInstitute = data.education[0]?.institute ?? "";

  const docChildren: Array<Paragraph | Table> = [];

  // ── Header table (Name/Role/Institute on left, Contact on right) ──
  const contactItems = [
    data.contact.phone ? `☎  ${data.contact.phone}` : null,
    data.contact.email ? `✉  ${data.contact.email}` : null,
    data.contact.linkedin ? `in  ${data.contact.linkedin}` : null,
    data.contact.website ? `◉  ${data.contact.website}` : null,
  ].filter(Boolean) as string[];

  const contactParagraphs = contactItems.map(
    (item) =>
      new Paragraph({
        children: [new TextRun({ text: item, size: 18, font: FONT, color: "374151" })],
        spacing: { after: 40 },
      })
  );

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          // Left: Name, role, institute
          new TableCell({
            width: { size: 60, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: data.name, bold: true, size: 40, font: FONT }),
                ],
                spacing: { after: 40 },
              }),
              ...(currentRole
                ? [
                    new Paragraph({
                      children: [
                        new TextRun({ text: currentRole, size: 22, font: FONT, color: "374151" }),
                      ],
                      spacing: { after: 20 },
                    }),
                  ]
                : []),
              ...(currentInstitute
                ? [
                    new Paragraph({
                      children: [
                        new TextRun({ text: currentInstitute, size: 22, font: FONT, color: "6B7280" }),
                      ],
                    }),
                  ]
                : []),
            ],
          }),
          // Right: Contact details
          new TableCell({
            width: { size: 40, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.TOP,
            children: contactParagraphs.length ? contactParagraphs : [new Paragraph("")],
          }),
        ],
      }),
    ],
  });

  docChildren.push(headerTable);

  // ── Summary ──
  if (data.summary) {
    docChildren.push(sectionHeader("PROFILE SUMMARY"));
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: data.summary, size: 20, font: FONT, italics: true })],
        spacing: { after: 80 },
      })
    );
  }

  // ── Experience ──
  if (data.experience.length) {
    docChildren.push(sectionHeader("WORK EXPERIENCE"));
    docChildren.push(experienceTable(data.experience));
  }

  // ── Projects ──
  if (data.projects.length) {
    docChildren.push(sectionHeader("PROJECTS"));
    for (const proj of data.projects) {
      docChildren.push(
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

  // ── Internships ──
  if (data.internships.length) {
    docChildren.push(sectionHeader("INTERNSHIPS"));
    docChildren.push(experienceTable(data.internships));
  }

  // ── Education (3-column table) ──
  if (data.education.length) {
    docChildren.push(sectionHeader("EDUCATION"));

    const eduTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: NO_BORDERS,
      rows: [
        // Header row
        new TableRow({
          children: ["DEGREE", "INSTITUTION", "YEAR"].map(
            (label) =>
              new TableCell({
                shading: { type: ShadingType.SOLID, color: LIGHT_GRAY, fill: LIGHT_GRAY },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: label, bold: true, size: 20, font: FONT })],
                  }),
                ],
              })
          ),
        }),
        // Data rows
        ...data.education.map(
          (edu) =>
            new TableRow({
              children: [edu.degree, edu.institute, edu.year].map(
                (val) =>
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: val, size: 20, font: FONT })],
                        spacing: { after: 40 },
                      }),
                    ],
                  })
              ),
            })
        ),
      ],
    });

    docChildren.push(eduTable);
  }

  // ── Achievements (2-column table) ──
  if (data.achievements.length) {
    docChildren.push(sectionHeader("ACHIEVEMENTS"));

    const achTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: NO_BORDERS,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.SOLID, color: LIGHT_GRAY, fill: LIGHT_GRAY },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "ACHIEVEMENT", bold: true, size: 20, font: FONT })],
                }),
              ],
            }),
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              shading: { type: ShadingType.SOLID, color: LIGHT_GRAY, fill: LIGHT_GRAY },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: "YEAR", bold: true, size: 20, font: FONT })],
                }),
              ],
            }),
          ],
        }),
        ...data.achievements.map(
          (ach) =>
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: ach.achievement, size: 20, font: FONT })],
                      spacing: { after: 40 },
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 15, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [new TextRun({ text: ach.year, size: 20, font: FONT })],
                    }),
                  ],
                }),
              ],
            })
        ),
      ],
    });

    docChildren.push(achTable);
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        children: docChildren,
      },
    ],
  });

  return Packer.toBlob(doc);
}
