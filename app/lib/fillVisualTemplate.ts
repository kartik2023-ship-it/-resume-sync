import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import type { Experience, StructuredResume } from "./types";

function esc(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Safe find-and-replace for literal strings in XML
function sub(xml: string, oldStr: string, newStr: string): string {
  return xml.split(oldStr).join(newStr);
}

// Replace all <w:t> text runs in a paragraph with a single new run (keeping pPr)
function replaceParagraphContent(paraXml: string, rPrXml: string, newText: string): string {
  const pTagMatch = paraXml.match(/^(<w:p[^>]*>)/);
  const pTag = pTagMatch ? pTagMatch[1] : "<w:p>";
  const pPrMatch = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : "";
  return `${pTag}${pPr}<w:r>${rPrXml}<w:t xml:space="preserve">${esc(newText)}</w:t></w:r></w:p>`;
}

// ── Table row factories ──

const TBL_PR = `<w:tblPr><w:tblW w:w="10643" w:type="dxa"/><w:tblInd w:w="130" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`;

const TBL_GRID = `<w:tblGrid><w:gridCol w:w="417"/><w:gridCol w:w="1579"/><w:gridCol w:w="1765"/><w:gridCol w:w="3958"/><w:gridCol w:w="1077"/><w:gridCol w:w="925"/><w:gridCol w:w="922"/></w:tblGrid>`;

// Dark section header row (gridSpan=7, fill=0C0C0C, white bold text)
function rowSectionHeader(title: string): string {
  return `<w:tr><w:trPr><w:trHeight w:val="244"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="10643" w:type="dxa"/><w:gridSpan w:val="7"/><w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:right w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="0C0C0C"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="1" w:line="223" w:lineRule="exact"/><w:ind w:left="105"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr><w:t>${esc(title)}</w:t></w:r></w:p></w:tc></w:tr>`;
}

// Work experience company header row
// Structure: [Full-Time rotated vMerge restart | company (gs=2) | role | dates (gs=3)]
function rowExpHeader(company: string, role: string, dates: string): string {
  return `<w:tr><w:trPr><w:trHeight w:val="246"/></w:trPr>` +
    // Cell 1: Full-Time label (col 1, rotated, vMerge restart)
    `<w:tc><w:tcPr><w:tcW w:w="417" w:type="dxa"/><w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>Full-Time</w:t></w:r></w:p></w:tc>` +
    // Cell 2: Company name (cols 2+3 gs=2)
    `<w:tc><w:tcPr><w:tcW w:w="3344" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="1" w:line="225" w:lineRule="exact"/><w:ind w:left="186"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>${esc(company.toUpperCase())}</w:t></w:r></w:p></w:tc>` +
    // Cell 3: Role (col 4)
    `<w:tc><w:tcPr><w:tcW w:w="3958" w:type="dxa"/><w:tcBorders><w:left w:val="nil"/><w:right w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="15" w:line="211" w:lineRule="exact"/><w:ind w:left="947"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${esc(role)} </w:t></w:r></w:p></w:tc>` +
    // Cell 4: Dates (cols 5+6+7 gs=3)
    `<w:tc><w:tcPr><w:tcW w:w="2924" w:type="dxa"/><w:gridSpan w:val="3"/><w:tcBorders><w:left w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="15" w:line="211" w:lineRule="exact"/><w:ind w:left="1213"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${esc(dates)}</w:t></w:r></w:p></w:tc>` +
    `</w:tr>`;
}

// Category + bullets row (vMerge continue for Full-Time column)
function rowExpCategory(categoryLabel: string, bullets: string[]): string {
  // Split category label across two paragraphs if it's long (template style)
  const words = categoryLabel.split(/\s+/);
  const half = Math.ceil(words.length / 2);
  const line1 = words.slice(0, half).join(" ");
  const line2 = words.slice(half).join(" ");

  const catParas =
    `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="1" w:line="225" w:lineRule="exact"/><w:ind w:left="-122" w:right="-145"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>${esc(line1)}</w:t></w:r></w:p>` +
    (line2
      ? `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="1" w:line="225" w:lineRule="exact"/><w:ind w:left="-122" w:right="-145"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>${esc(line2)}</w:t></w:r></w:p>`
      : "");

  const bulletParas = bullets
    .map(
      (b) =>
        `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr><w:ind w:left="286" w:hanging="142"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(b)}</w:t></w:r></w:p>`
    )
    .join("");

  return (
    `<w:tr><w:trPr><w:trHeight w:val="246"/></w:trPr>` +
    // Cell 1: vMerge continue
    `<w:tc><w:tcPr><w:tcW w:w="417" w:type="dxa"/><w:vMerge/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:tc>` +
    // Cell 2: Category label (col 2)
    `<w:tc><w:tcPr><w:tcW w:w="1579" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:vAlign w:val="center"/></w:tcPr>${catParas}</w:tc>` +
    // Cell 3: Bullets (cols 3-7 gs=5)
    `<w:tc><w:tcPr><w:tcW w:w="8647" w:type="dxa"/><w:gridSpan w:val="5"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:vAlign w:val="center"/></w:tcPr>${bulletParas}</w:tc>` +
    `</w:tr>`
  );
}

// Projects row — all projects in a single wide cell
function rowProjects(projects: { name: string; description: string }[]): string {
  const paras = projects
    .map(
      (p) =>
        `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:line="276" w:lineRule="auto"/><w:ind w:left="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(p.name)}</w:t></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">  —  ${esc(p.description)}</w:t></w:r></w:p>`
    )
    .join("");
  return `<w:tr><w:trPr><w:trHeight w:val="320"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="10643" w:type="dxa"/><w:gridSpan w:val="7"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr>${paras}</w:tc></w:tr>`;
}

// Internship header row — same structure as experience but label "Intern"
function rowInternHeader(type: string, company: string, role: string, dates: string): string {
  return (
    `<w:tr><w:trPr><w:trHeight w:val="220"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="417" w:type="dxa"/><w:vMerge w:val="restart"/><w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:ind w:left="395"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>${esc(type)}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="3344" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:right w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:line="200" w:lineRule="exact"/><w:ind w:left="273"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>${esc("  " + company.toUpperCase())}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="3958" w:type="dxa"/><w:tcBorders><w:left w:val="nil"/><w:right w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>${esc(role)}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2924" w:type="dxa"/><w:gridSpan w:val="3"/><w:tcBorders><w:left w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>${esc(dates)}</w:t></w:r></w:p></w:tc>` +
    `</w:tr>`
  );
}

// Internship category+bullets row (same structure as exp category but different shade)
function rowInternCategory(categoryLabel: string, bullets: string[]): string {
  const words = categoryLabel.split(/\s+/);
  const half = Math.ceil(words.length / 2);
  const line1 = words.slice(0, half).join(" ");
  const line2 = words.slice(half).join(" ");

  const catParas =
    `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:ind w:left="-122" w:right="-145"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>${esc(line1)}</w:t></w:r></w:p>` +
    (line2
      ? `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:ind w:left="-122" w:right="-145"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>${esc(line2)}</w:t></w:r></w:p>`
      : "");

  const bulletParas = bullets
    .map(
      (b) =>
        `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr><w:ind w:left="286" w:hanging="142"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(b)}</w:t></w:r></w:p>`
    )
    .join("");

  return (
    `<w:tr><w:trPr><w:trHeight w:val="246"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="417" w:type="dxa"/><w:vMerge/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="1579" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:vAlign w:val="center"/></w:tcPr>${catParas}</w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="8647" w:type="dxa"/><w:gridSpan w:val="5"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:vAlign w:val="center"/></w:tcPr>${bulletParas}</w:tc>` +
    `</w:tr>`
  );
}

// Education header row (col 1+2+3 | col 3+4+5 | col 7)
function rowEduHeader(): string {
  return (
    `<w:tr><w:trPr><w:trHeight w:val="227"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="3761" w:type="dxa"/><w:gridSpan w:val="3"/><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="15" w:line="192" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>Degree/Exam</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="5960" w:type="dxa"/><w:gridSpan w:val="3"/><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:spacing w:before="15" w:line="192" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>Board/Institute</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="922" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D6D6D6"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>Year</w:t></w:r></w:p></w:tc>` +
    `</w:tr>`
  );
}

function rowEduData(degree: string, institute: string, year: string): string {
  return (
    `<w:tr>` +
    `<w:tc><w:tcPr><w:tcW w:w="3761" w:type="dxa"/><w:gridSpan w:val="3"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(degree)}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="5960" w:type="dxa"/><w:gridSpan w:val="3"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(institute)}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="922" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>${esc(year)}</w:t></w:r></w:p></w:tc>` +
    `</w:tr>`
  );
}

// Achievements row — "Achievement" label cell + bullets cell
function rowAchievements(achievements: { achievement: string; year: string }[]): string {
  const bulletParas = achievements
    .map((a) => {
      const text = a.year ? `${a.achievement}  (${a.year})` : a.achievement;
      return `<w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr><w:spacing w:before="10"/><w:ind w:left="240" w:hanging="142"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
    })
    .join("");
  return (
    `<w:tr><w:trPr><w:trHeight w:val="441"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="1996" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableParagraph"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>Achievement</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="8647" w:type="dxa"/><w:gridSpan w:val="5"/></w:tcPr>${bulletParas}</w:tc>` +
    `</w:tr>`
  );
}

// Build the full table
function buildTable(data: StructuredResume): string {
  const rows: string[] = [];

  // ── Work Experience section ──
  rows.push(rowSectionHeader("WORK EXPERIENCE"));
  for (const exp of data.experience) {
    rows.push(rowExpHeader(exp.company, exp.role, exp.dates));
    const cats = exp.categories.length > 0 ? exp.categories : [{ label: "Responsibilities", bullets: [] }];
    for (const cat of cats) {
      rows.push(rowExpCategory(cat.label || "Responsibilities", cat.bullets));
    }
  }

  // ── Projects ──
  if (data.projects.length) {
    rows.push(rowSectionHeader("PROJECTS & PORTFOLIO"));
    rows.push(rowProjects(data.projects));
  }

  // ── Internships ──
  if (data.internships.length) {
    rows.push(rowSectionHeader("INTERNSHIP & PROJECTS"));
    for (const intern of data.internships) {
      rows.push(rowInternHeader("Intern", intern.company, intern.role, intern.dates));
      for (const cat of intern.categories) {
        rows.push(rowInternCategory(cat.label || "Responsibilities", cat.bullets));
      }
    }
  }

  // ── Education ──
  if (data.education.length) {
    rows.push(rowSectionHeader("ACADEMIC QUALIFICATIONS"));
    rows.push(rowEduHeader());
    for (const edu of data.education) {
      rows.push(rowEduData(edu.degree, edu.institute, edu.year));
    }
  }

  // ── Achievements ──
  if (data.achievements.length) {
    rows.push(rowSectionHeader("ACHIEVEMENTS"));
    rows.push(rowAchievements(data.achievements));
  }

  return `<w:tbl>${TBL_PR}${TBL_GRID}${rows.join("")}</w:tbl>`;
}

export function fillVisualTemplate(data: StructuredResume): Buffer {
  const templatePath = path.join(process.cwd(), "templates", "visual-template.docx");
  const templateBuffer = fs.readFileSync(templatePath);
  const zip = new PizZip(templateBuffer);

  let xml = zip.file("word/document.xml")!.asText();

  // Extract document wrapper and sectPr
  const bodyStart = xml.indexOf("<w:body>");
  const bodyEnd = xml.lastIndexOf("</w:body>") + "</w:body>".length;
  const before = xml.slice(0, bodyStart);
  const after = xml.slice(bodyEnd);

  const sectPrMatch = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : "";

  // Extract the pre-table area and the main table
  const tableStart = xml.indexOf("<w:tbl>");
  const tableEnd = xml.lastIndexOf("</w:tbl>") + "</w:tbl>".length;
  let preTableXml = xml.slice(xml.indexOf("<w:body>") + "<w:body>".length, tableStart);

  // ── Text replacements in pre-table area ──

  // 1. Name ("Kartik" appears in body paragraphs — replace all occurrences before table)
  preTableXml = sub(preTableXml, ">Kartik<", `>${esc(data.name)}<`);

  // 2. Contact info — replace known literal strings
  if (data.contact.email) {
    preTableXml = sub(preTableXml, "kartik2023@email.iimcal.ac.in", esc(data.contact.email));
  }
  if (data.contact.linkedin) {
    // template has "linkedin.com/in/1996kartik/"
    preTableXml = preTableXml.replace(
      /linkedin\.com\/in\/[^<"]+/g,
      esc(data.contact.linkedin.replace(/^https?:\/\//i, ""))
    );
  }
  if (data.contact.website) {
    preTableXml = preTableXml.replace(
      /https?:\/\/kartikaipm[^<"]+/gi,
      esc(data.contact.website)
    );
  }

  // 3. Role + company paragraph (Para 4):
  //    "Assistant Product Manager, " + "Policybazaar" → new role + ", " + new company
  const currentRole = data.experience[0]?.role ?? "";
  const currentCompany = data.experience[0]?.company ?? "";
  if (currentRole && currentCompany) {
    // The role para has pPr with spacing/indent — keep it, replace the runs
    preTableXml = preTableXml.replace(
      /(<w:p[^>]*>)(<w:pPr><w:spacing w:before="29"[^<]*<\/w:spacing><w:ind w:left="268"[^/]*\/>[\s\S]*?<\/w:pPr>)[\s\S]*?(<\/w:p>)/,
      (_, pTag, pPr, pClose) =>
        `${pTag}${pPr}<w:r><w:rPr><w:rFonts w:ascii="Cambria"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(currentRole)}, ${esc(currentCompany)}</w:t></w:r>${pClose}`
    );
  }

  // 4. Institute paragraph (Para 5 has drawing + institute text run):
  //    Replace "Indian Institute of Management, Calcutta (2021-23)"
  const currentEdu = data.education[0];
  if (currentEdu) {
    preTableXml = preTableXml.replace(
      /Indian Institute of Management[^<]*/g,
      esc(currentEdu.institute + (currentEdu.year ? ` (${currentEdu.year})` : ""))
    );
  }

  // 5. Summary paragraph (Para 6 — italic, multi-run):
  //    Replace the entire content with the new summary
  if (data.summary) {
    preTableXml = preTableXml.replace(
      /(<w:p[^>]*><w:pPr><w:spacing w:before="4"\/>[\s\S]*?<\/w:pPr>)[\s\S]*?(<\/w:p>)/,
      (_, pPrPart, pClose) =>
        `${pPrPart}<w:r><w:rPr><w:i/><w:iCs/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">${esc(data.summary)}</w:t></w:r>${pClose}`
    );
  }

  // ── Rebuild the main table ──
  const newTable = buildTable(data);

  // Reconstruct full XML
  const newXml =
    before +
    `<w:body>${preTableXml}${newTable}${sectPr}</w:body>` +
    after;

  zip.file("word/document.xml", newXml);
  return Buffer.from(zip.generate({ type: "arraybuffer" }));
}
