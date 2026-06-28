export interface ResumeContact {
  phone: string;
  email: string;
  linkedin: string;
  website: string;
}

export interface ExperienceCategory {
  label: string;
  bullets: string[];
}

export interface Experience {
  company: string;
  role: string;
  dates: string;
  categories: ExperienceCategory[];
}

export interface Project {
  name: string;
  description: string;
}

export interface Education {
  degree: string;
  institute: string;
  year: string;
}

export interface Achievement {
  achievement: string;
  year: string;
}

export interface StructuredResume {
  name: string;
  contact: ResumeContact;
  summary: string;
  experience: Experience[];
  projects: Project[];
  internships: Experience[];
  education: Education[];
  achievements: Achievement[];
}

export interface ATSScore {
  score: number;
  keywordMatch: number;
  skillsCoverage: number;
  suggestions: string[];
}

export interface RewriteResponse {
  resume: string;
  structured: StructuredResume;
  ats: ATSScore;
}
