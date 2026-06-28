import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import type { ResumeContact, StructuredResume } from "./types";

function esc(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function pName(name: string): string {
  return `<w:p><w:pPr><w:spacing w:after="80"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr><w:t>${esc(name)}</w:t></w:r></w:p>`;
}

function pContact(c: ResumeContact): string {
  const parts = [c.phone, c.email, c.linkedin, c.website].filter(Boolean);
  const runs = parts
    .map((v, i) => {
      const val = `<w:r><w:rPr><w:color w:val="444444"/><w:sz w:val="18"/><w:szCs w:val="16"/></w:rPr><w:t>${esc(v)}</w:t></w:r>`;
      const sep =
        i < parts.length - 1
          ? `<w:r><w:rPr><w:color w:val="999999"/><w:sz w:val="18"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">   |   </w:t></w:r>`
          : "";
      return val + sep;
    })
    .join("");
  return `<w:p><w:pPr><w:spacing w:after="40"/><w:jc w:val="center"/></w:pPr>${runs}</w:p>`;
}

function pSummary(text: string): string {
  return `<w:p><w:r><w:rPr><w:color w:val="222222"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

// experience/internship header: role  —  company  [tab]  dates
function pEntryHeader(role: string, company: string, dates: string, spacingAfter = "20"): string {
  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="10080"/></w:tabs><w:spacing w:before="160" w:after="${spacingAfter}"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">${esc(role)}</w:t></w:r><w:r><w:rPr><w:color w:val="777777"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">  —  </w:t></w:r><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">${esc(company)}</w:t></w:r><w:r><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:i/><w:iCs/><w:color w:val="555555"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(dates)}</w:t></w:r></w:p>`;
}

function pCategory(label: string): string {
  return `<w:p><w:pPr><w:spacing w:before="100" w:after="20"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="333333"/><w:sz w:val="19"/><w:szCs w:val="19"/><w:u w:val="single" w:color="AAAAAA"/></w:rPr><w:t>${esc(label)}</w:t></w:r></w:p>`;
}

function pBullet(text: string): string {
  return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:spacing w:before="30" w:after="30"/></w:pPr><w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

function pProject(name: string, description: string): string {
  return `<w:p><w:pPr><w:spacing w:before="100" w:after="20"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(name)}</w:t></w:r><w:r><w:rPr><w:color w:val="999999"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">  —  </w:t></w:r><w:r><w:rPr><w:color w:val="444444"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(description)}</w:t></w:r></w:p>`;
}

function pEdu(degree: string, institute: string, year: string): string {
  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="10080"/></w:tabs><w:spacing w:before="100" w:after="30"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">${esc(degree)}  —  ${esc(institute)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:i/><w:iCs/><w:color w:val="555555"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(year)}</w:t></w:r></w:p>`;
}

function pAchievement(text: string): string {
  return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:spacing w:before="30" w:after="30"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

export function fillATSTemplate(data: StructuredResume): Buffer {
  const templatePath = path.join(process.cwd(), "templates", "ats-template.docx");
  const templateBuffer = fs.readFileSync(templatePath);
  const zip = new PizZip(templateBuffer);

  const xml = zip.file("word/document.xml")!.asText();

  // Extract all template paragraphs (used as-is for static sections)
  const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];

  // Extract sectPr (page setup: margins, size, etc.)
  const sectPrMatch = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : "";

  // Get document wrapper (everything outside <w:body>)
  const bodyStart = xml.indexOf("<w:body>");
  const bodyEnd = xml.lastIndexOf("</w:body>") + "</w:body>".length;
  const before = xml.slice(0, bodyStart);
  const after = xml.slice(bodyEnd);

  const out: string[] = [];

  // ── Header ──
  out.push(pName(data.name));
  out.push(pContact(data.contact));
  out.push(paras[2]); // empty paragraph

  // ── Professional Summary ──
  out.push(paras[3]); // section header
  out.push(paras[4]); // empty
  if (data.summary) out.push(pSummary(data.summary));

  // ── Core Skills (keep template content unchanged) ──
  out.push(paras[6]); // CORE SKILLS header
  out.push(paras[7]); // empty
  out.push(paras[8], paras[9], paras[10], paras[11]); // 4 skill lines

  // ── Work Experience ──
  out.push(paras[12]); // WORK EXPERIENCE header
  for (const exp of data.experience) {
    out.push(pEntryHeader(exp.role, exp.company, exp.dates, "20"));
    const multi = exp.categories.length > 1;
    for (const cat of exp.categories) {
      if (cat.label && multi) out.push(pCategory(cat.label));
      for (const b of cat.bullets) out.push(pBullet(b));
    }
  }

  // ── Projects ──
  if (data.projects.length) {
    out.push(paras[37]); // PROJECTS & PORTFOLIO header
    for (const proj of data.projects) out.push(pProject(proj.name, proj.description));
  }

  // ── Internships ──
  if (data.internships.length) {
    out.push(paras[43]); // INTERNSHIPS header
    for (const intern of data.internships) {
      out.push(pEntryHeader(intern.role, intern.company, intern.dates, "30"));
      for (const cat of intern.categories) {
        for (const b of cat.bullets) out.push(pBullet(b));
      }
    }
  }

  // ── Education ──
  if (data.education.length) {
    out.push(paras[50]); // EDUCATION header
    for (const edu of data.education) out.push(pEdu(edu.degree, edu.institute, edu.year));
  }

  // ── Achievements ──
  if (data.achievements.length) {
    out.push(paras[55]); // ACHIEVEMENTS & CERTIFICATIONS header
    for (const ach of data.achievements) {
      const text = ach.year ? `${ach.achievement}  (${ach.year})` : ach.achievement;
      out.push(pAchievement(text));
    }
  }

  // Rebuild body
  const newXml =
    before +
    `<w:body>${out.join("")}${sectPr}</w:body>` +
    after;

  zip.file("word/document.xml", newXml);
  return Buffer.from(zip.generate({ type: "arraybuffer" }));
}
