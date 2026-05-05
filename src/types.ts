
export type DataSource = 'yellow' | 'white' | 'green' | 'blue' | 'purple';

export interface WorkloadRow {
  id: string;
  disciplineName: string; // Yellow
  educationalProgram: string; // Yellow
  credits: number; // Yellow
  language: string; // White
  course: number; // Yellow
  semester: number; // Yellow
  groupCount: number; // White
  subGroupCount: number; // White
  studentCount: number; // White
  groupNumber: string; // White
  
  // Lectures
  lecturePlan: number; // Yellow
  lectureTotal: number; // Green (Calculated)
  
  // Laboratory
  labPlan: number; // Yellow
  labTotal: number; // Green (Calculated)
  labCredits: number; // Green (Calculated)
  
  srop: number; // White
  totalCredits: number; // Green (Calculated)
  
  courseWork: boolean; // Yellow
  courseWorkCredits: number; // Green (Calculated)
  
  ppsName: string; // Blue/Purple (Faculty member)
}

export interface SignatureAuthority {
  role: string;
  name: string;
}

export const AUTHORITIES: SignatureAuthority[] = [
  { role: "Директор ДАВ", name: "Жургенов Ж.С." },
  { role: "Зам. директора ДАВ", name: "Нургалиева А.Ш." },
  { role: "Декан факультета КСИПО", name: "—" },
  { role: "Заместитель декана КСИПО", name: "Жетписбаева А.Е." },
  { role: "И.о. зав. кафедрой 'Компьютерные науки'", name: "Аканова A.C." }
];
