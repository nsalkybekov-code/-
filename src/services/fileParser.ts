
import * as XLSX from 'xlsx';
import { WorkloadRow } from '../types';

export async function parseExcelFile(file: File): Promise<Partial<WorkloadRow>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Преобразуем в массив массивов для гибкого поиска заголовков
        const rows_raw = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        if (rows_raw.length === 0) return resolve([]);

        // Ищем строку с заголовками более тщательно
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(rows_raw.length, 30); i++) {
          const rowText = rows_raw[i].map(c => String(c || '').toLowerCase()).join(' ');
          // Ищем строку, где есть и дисциплина, и кредиты/курс
          if ((rowText.includes('дисциплин') || rowText.includes('предмет')) && 
              (rowText.includes('кред') || rowText.includes('курс') || rowText.includes('сем'))) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          headerRowIndex = 0; 
        }

        const headers = Array.from(rows_raw[headerRowIndex] || []).map(h => String(h || '').toLowerCase().trim());
        const dataRows = rows_raw.slice(headerRowIndex + 1);

        const findIdx = (keywords: string[], exact = false) => {
          if (exact) return headers.findIndex(h => keywords.some(k => h === k));
          // Сначала пробуем найти максимально близкое к началу или длинное совпадение
          return headers.findIndex(h => h && keywords.some(k => h.includes(k)));
        };

        const idxMap = {
          name: findIdx(['дисциплин', 'наименование', 'предмет', 'название']),
          credits: findIdx(['кредит', 'кред', 'объем', 'ects']),
          course: findIdx(['курс']),
          semester: findIdx(['семестр', 'сем']),
          op: findIdx(['образовательная программа', 'оп', 'шифр', 'специальность']),
          lecture: findIdx(['лекция', 'лек.', 'лекц']),
          lab: findIdx(['лабор', 'лаб.', 'лаб']),
          courseWork: findIdx(['курсовая', 'кр', 'к.р.', 'к/р']),
          language: findIdx(['язык', 'отделение']),
          students: findIdx(['студ', 'количество студентов', 'контингент', 'кол-во']),
          groups: findIdx(['групп']),
          pps: findIdx(['фио', 'преподаватель', 'лектор', 'ппс']),
        };

        const getNum = (val: any) => {
          if (typeof val === 'number') return val;
          const str = String(val || '').replace(',', '.').replace(/[^\d.]/g, '');
          const n = parseFloat(str);
          return isNaN(n) ? 0 : n;
        };

        const result = dataRows
          .filter(row => {
            if (idxMap.name === -1) return false;
            const val = String(row[idxMap.name] || '').trim();
            if (!val || val.length < 3) return false;
            
            const lowerVal = val.toLowerCase();
            
            // 1. Игнорируем итоговые строки и технические подзаголовки
            if (lowerVal.includes('итого') || lowerVal.includes('всего') || lowerVal === 'дисциплина' || lowerVal === 'наименование') return false;
            
            // 2. Дисциплины-исключения, которые мы берем ВСЕГДА (защита от ложного срабатывания фильтра заголовков)
            const isTarget = lowerVal.includes('программ') || lowerVal.includes('основ') || lowerVal.includes('информ');

            // 3. Проверка на заголовок цикла или раздела
            const cycleKeywords = ['цикл', 'компонент', 'гуманитарно', 'социальн', 'базов', 'профилир', 'естественно', 'общеобразовател', 'элективн', 'вузовск'];
            const isKnownCycle = cycleKeywords.some(k => lowerVal.includes(k));
            const isNumericHeader = /^\d+\.?\s+[А-Я]/.test(val) && val.split(' ').length < 5;

            if ((isKnownCycle || isNumericHeader) && !isTarget) {
               return false;
            }

            // 4. Проверка на наличие данных (кредиты или студенты). 
            // Если в строке нет ни кредитов, ни студентов, и это не спец-цель, то это скорее всего мусорная строка.
            const credits = idxMap.credits !== -1 ? getNum(row[idxMap.credits]) : 0;
            const students = idxMap.students !== -1 ? getNum(row[idxMap.students]) : 0;
            const hasData = credits > 0 || students > 0;
            
            return hasData || isTarget;
          }) 
          .map((row) => ({
            disciplineName: idxMap.name !== -1 ? String(row[idxMap.name] || '').trim() : '',
            credits: idxMap.credits !== -1 ? getNum(row[idxMap.credits]) : 0,
            course: idxMap.course !== -1 ? getNum(row[idxMap.course]) : 1,
            semester: idxMap.semester !== -1 ? getNum(row[idxMap.semester]) : 1,
            educationalProgram: idxMap.op !== -1 ? String(row[idxMap.op] || '').trim() : '',
            lecturePlan: idxMap.lecture !== -1 ? getNum(row[idxMap.lecture]) : 0,
            labPlan: idxMap.lab !== -1 ? getNum(row[idxMap.lab]) : 0,
            courseWork: !!(idxMap.courseWork !== -1 && row[idxMap.courseWork] && 
                         (String(row[idxMap.courseWork]).toLowerCase().includes('да') || 
                          String(row[idxMap.courseWork]) === '+')),
            language: idxMap.language !== -1 ? String(row[idxMap.language] || 'Русский').trim() : 'Русский',
            studentCount: idxMap.students !== -1 ? getNum(row[idxMap.students]) : 0,
            groupCount: idxMap.groups !== -1 ? getNum(row[idxMap.groups]) : 1,
            ppsName: idxMap.pps !== -1 ? String(row[idxMap.pps] || '').trim() : '',
          }));
        
        resolve(result);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export async function parsePDFFile(file: File): Promise<Partial<WorkloadRow>[]> {
  // PDF parsing is typically more complex and often requires OCR or specific structured positioning
  // For this demonstration, we'll return a placeholder success message or mock data 
  // until a robust PDF parsing strategy is implemented.
  console.log("PDF Parsing initiated for:", file.name);
  return [];
}
