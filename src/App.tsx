/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useCallback } from 'react';
import { 
  FileText, 
  Upload, 
  Plus, 
  Trash2, 
  Download, 
  CheckCircle2, 
  AlertCircle,
  Users,
  BookOpen,
  Calculator
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { WorkloadRow, AUTHORITIES } from './types.ts';
import { parseExcelFile } from './services/fileParser.ts';
import { cn } from './lib/utils.ts';

export default function App() {
  const [rows, setRows] = useState<WorkloadRow[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Initial empty row template
  const createEmptyRow = useCallback((): WorkloadRow => ({
    id: Math.random().toString(36).substr(2, 9),
    disciplineName: '',
    educationalProgram: '',
    credits: 0,
    language: 'Русский',
    course: 1,
    semester: 1,
    groupCount: 1,
    subGroupCount: 1,
    studentCount: 0,
    groupNumber: '',
    lecturePlan: 0,
    lectureTotal: 0,
    labPlan: 0,
    labTotal: 0,
    labCredits: 0,
    srop: 0,
    totalCredits: 0,
    courseWork: false,
    courseWorkCredits: 0,
    ppsName: '',
  }), []);

  // Centralized calculation logic
  const calculateRow = (row: WorkloadRow): WorkloadRow => {
    const lectureTotal = row.lecturePlan * row.groupCount;
    const labTotal = row.labPlan * row.subGroupCount;
    const labCredits = Number((labTotal / 30).toFixed(2));
    const courseWorkCredits = row.courseWork ? 1 : 0;
    
    const totalCredits = Number((
      row.credits + 
      labCredits + 
      courseWorkCredits
    ).toFixed(2));

    return {
      ...row,
      lectureTotal,
      labTotal,
      labCredits,
      courseWorkCredits,
      totalCredits
    };
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsUploading(true);
    setError(null);
    try {
      const allNewRows: WorkloadRow[] = [];
      
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const parsedRows = await parseExcelFile(file);
          const formattedRows = parsedRows.map(p => {
            const empty = createEmptyRow();
            const merged = { ...empty, ...p };
            return calculateRow(merged);
          });
          allNewRows.push(...formattedRows);
        } catch (fileErr) {
          console.error(`Ошибка в файле ${file.name}:`, fileErr);
        }
      }

      if (allNewRows.length > 0) {
        setRows(prev => {
          const combined = [...prev, ...allNewRows];
          return combined;
        });
      } else {
        setError('Не удалось извлечь данные из выбранных файлов.');
      }
    } catch (err) {
      setError('Ошибка при пакетной загрузке файлов.');
      console.error(err);
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  const updateRow = (id: string, updates: Partial<WorkloadRow>) => {
    setRows(prev => prev.map(row => {
      if (row.id !== id) return row;
      const updated = { ...row, ...updates };
      return calculateRow(updated);
    }));
  };

  const removeRow = (id: string) => {
    setRows(prev => prev.filter(r => r.id !== id));
  };

  const exportToExcel = () => {
    if (rows.length === 0) return;
    
    // Пакетная подготовка данных для экспорта
    const exportData = rows.map(r => ({
      'Дисциплина': r.disciplineName,
      'ОП': r.educationalProgram,
      'Кред (РУП)': r.credits,
      'Язык': r.language,
      'Курс': r.course,
      'Сем.': r.semester,
      'Групп': r.groupCount,
      'Подгрупп': r.subGroupCount,
      'Студ.': r.studentCount,
      '№ Группы': r.groupNumber,
      'Лек (План)': r.lecturePlan,
      'Лек (Всего)': r.lectureTotal,
      'Лаб (План)': r.labPlan,
      'Лаб (Всего)': r.labTotal,
      'Лаб (Кред)': r.labCredits,
      'СРОП': r.srop,
      'Всего кред': r.totalCredits,
      'Курсовая': r.courseWork ? 'Да' : 'Нет',
      'КР (Кред)': r.courseWorkCredits,
      'ППС': r.ppsName
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Нагрузка");

    // Генерируем имя файла с датой
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Academic_Workload_${date}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans">
      <header className="max-w-[1920px] mx-auto mb-8">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold text-slate-900 tracking-tight">Менеджер академической нагрузки</h1>
            <p className="text-slate-500 mt-1">Учет РУП и расчет часов</p>
          </div>
          
          <div className="flex items-center gap-3">
            <label className={cn(
              "flex items-center gap-2 px-6 py-3 text-white rounded-xl font-semibold transition-all cursor-pointer shadow-lg active:scale-95",
              isUploading ? "bg-slate-400 cursor-not-allowed" : "bg-indigo-600 hover:bg-indigo-700 hover:shadow-indigo-200"
            )}>
              <Upload size={20} />
              <span>{isUploading ? 'Загрузка...' : 'Загрузить (Excel/PDF)'}</span>
              <input 
                type="file" 
                className="hidden" 
                multiple 
                accept=".xlsx,.xls,.pdf" 
                onChange={handleFileUpload} 
                disabled={isUploading}
              />
            </label>
            <button 
              onClick={exportToExcel}
              disabled={rows.length === 0}
              className={cn(
                "flex items-center gap-2 px-6 py-3 text-white rounded-xl font-semibold transition-all shadow-lg active:scale-95",
                rows.length === 0 ? "bg-slate-300 cursor-not-allowed" : "bg-blue-600 hover:bg-blue-700 hover:shadow-blue-200"
              )}
            >
              <Download size={20} />
              <span>Скачать (Excel)</span>
            </button>
            <button 
              onClick={() => {
                const newRow = calculateRow(createEmptyRow());
                setRows(prev => [...prev, newRow]);
              }}
              className="flex items-center gap-2 px-6 py-3 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl font-semibold transition-all shadow-md active:scale-95"
            >
              <Plus size={20} />
              <span>Добавить строку</span>
            </button>
          </div>
        </div>

        {error && (
          <motion.div 
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            className="mt-4 p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl flex items-center gap-2 shadow-sm"
          >
            <AlertCircle size={20} />
            <span className="text-sm font-semibold">{error}</span>
          </motion.div>
        )}
      </header>

      <main className="max-w-[100vw] overflow-x-auto bg-white rounded-2xl shadow-2xl border border-slate-200 mx-auto">
        <div className="min-w-[2000px]">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr className="bg-slate-800 text-white border-b border-slate-700">
                <th rowSpan={2} className="p-4 border-r border-slate-700 sticky left-0 bg-slate-800 z-30 uppercase tracking-widest text-[9px] w-64 text-left">Дисциплина / ОП</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-20">Кред (РУП)</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-28">Язык</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-14">Курс</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-14">Сем.</th>
                <th colSpan={4} className="p-2 border-r border-slate-700 text-center bg-slate-700">Группы (Ввод)</th>
                <th colSpan={2} className="p-2 border-r border-slate-700 text-center bg-yellow-900/30">Лекции</th>
                <th colSpan={3} className="p-2 border-r border-slate-700 text-center bg-green-900/30">Лабораторные</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-16">СРОП</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center w-20">Всего кред.</th>
                <th colSpan={2} className="p-2 border-r border-slate-700 text-center bg-yellow-900/30">Курсовая</th>
                <th rowSpan={2} className="p-2 border-r border-slate-700 text-center">ППС (Штатный)</th>
                <th rowSpan={2} className="p-2 text-center w-12"></th>
              </tr>
              <tr className="bg-slate-700/50 text-white/80 border-b border-slate-600">
                <th className="p-2 border-r border-slate-600 text-[9px]">Групп</th>
                <th className="p-2 border-r border-slate-600 text-[9px]">Подгрупп</th>
                <th className="p-2 border-r border-slate-600 text-[9px]">Студ.</th>
                <th className="p-2 border-r border-slate-600 text-[9px]">№ группы</th>
                
                <th className="p-2 border-r border-slate-600 text-[9px] bg-yellow-500/20">План(Ж)</th>
                <th className="p-2 border-r border-slate-600 text-[9px] bg-green-500/20">Всего(З)</th>
                
                <th className="p-2 border-r border-slate-600 text-[9px] bg-yellow-500/20">План(Ж)</th>
                <th className="p-2 border-r border-slate-600 text-[9px] bg-green-500/20">Всего(З)</th>
                <th className="p-2 border-r border-slate-600 text-[9px] bg-green-500/20">Кред(З)</th>
                
                <th className="p-2 border-r border-slate-600 text-[9px] bg-yellow-500/20" title="Курсовая работа">КР?</th>
                <th className="p-2 border-r border-slate-600 text-[9px] bg-green-500/20">Кред(З)</th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr>
                  <td colSpan={20} className="p-20 text-center text-slate-400 bg-slate-50/50">
                    <div className="flex flex-col items-center gap-4">
                      <div className="p-6 bg-white rounded-full shadow-lg">
                        <BookOpen size={48} className="text-slate-200" />
                      </div>
                      <p className="text-lg font-medium text-slate-500">Нет данных. Загрузите РУП или добавьте строку вручную.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                <AnimatePresence>
                  {rows.map((row) => (
                    <motion.tr 
                      key={row.id}
                      initial={{ opacity: 0, x: -20 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, height: 0 }}
                      className="border-b border-slate-200 hover:bg-slate-50/50 group"
                    >
                      {/* Discipline - Yellow */}
                      <td className="p-2 border-r border-slate-200 sticky left-0 bg-white group-hover:bg-slate-50 z-20">
                        <textarea
                          placeholder="Наименование дисциплины"
                          value={row.disciplineName}
                          onChange={(e) => updateRow(row.id, { disciplineName: e.target.value })}
                          className="w-full bg-yellow-50/80 border border-yellow-200 rounded-md p-2 text-[11px] leading-tight resize-none h-16 focus:ring-2 ring-yellow-400 focus:outline-none transition-all shadow-inner"
                        />
                        <input
                          placeholder="Образовательная программа (ОП)"
                          value={row.educationalProgram}
                          onChange={(e) => updateRow(row.id, { educationalProgram: e.target.value })}
                          className="w-full mt-1.5 bg-yellow-50/50 border border-yellow-100 rounded p-1.5 text-[9px] font-medium text-yellow-800"
                        />
                      </td>

                      {/* Credits (RUP) - Yellow */}
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/30">
                        <input
                          type="number"
                          value={row.credits || ''}
                          onChange={(e) => updateRow(row.id, { credits: Number(e.target.value) })}
                          className="w-full h-12 bg-transparent text-center font-bold text-slate-700 focus:outline-none"
                        />
                      </td>

                      {/* Language - White */}
                      <td className="p-1 border-r border-slate-200">
                        <select
                          value={row.language}
                          onChange={(e) => updateRow(row.id, { language: e.target.value })}
                          className="w-full h-12 bg-white text-center rounded focus:outline-none text-[11px]"
                        >
                          <option>Русский</option>
                          <option>Казахский</option>
                          <option>Английский</option>
                        </select>
                      </td>

                      {/* Course/Semester - Yellow */}
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/30 text-center">
                        <input type="number" value={row.course || ''} onChange={(e) => updateRow(row.id, { course: Number(e.target.value) })} className="w-full bg-transparent text-center focus:outline-none font-medium" />
                      </td>
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/30 text-center">
                        <input type="number" value={row.semester || ''} onChange={(e) => updateRow(row.id, { semester: Number(e.target.value) })} className="w-full bg-transparent text-center focus:outline-none font-medium" />
                      </td>

                      {/* Groups - White (Manual) */}
                      <td className="p-1 border-r border-slate-200">
                        <input
                          type="number"
                          value={row.groupCount || ''}
                          onChange={(e) => updateRow(row.id, { groupCount: Number(e.target.value) })}
                          className="w-full h-12 bg-white text-center font-bold text-slate-800 border-2 border-slate-100 rounded-md focus:border-indigo-400 focus:outline-none"
                        />
                      </td>
                      <td className="p-1 border-r border-slate-200">
                        <input
                          type="number"
                          value={row.subGroupCount || ''}
                          onChange={(e) => updateRow(row.id, { subGroupCount: Number(e.target.value) })}
                          className="w-full h-12 bg-white text-center font-bold text-slate-800 border-2 border-slate-100 rounded-md focus:border-indigo-400 focus:outline-none"
                        />
                      </td>
                      <td className="p-1 border-r border-slate-200">
                        <input
                          type="number"
                          value={row.studentCount || ''}
                          onChange={(e) => updateRow(row.id, { studentCount: Number(e.target.value) })}
                          className="w-full h-12 bg-white text-center font-bold text-slate-800 border-2 border-slate-100 rounded-md focus:border-indigo-400 focus:outline-none"
                        />
                      </td>
                      <td className="p-1 border-r border-slate-200">
                        <input
                          type="text"
                          value={row.groupNumber || ''}
                          onChange={(e) => updateRow(row.id, { groupNumber: e.target.value })}
                          className="w-full h-12 bg-white text-center border-2 border-slate-100 rounded-md focus:border-indigo-400 focus:outline-none font-mono"
                          placeholder="№"
                        />
                      </td>

                      {/* Lectures - Plan(Yellow), Total(Green) */}
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/50">
                        <input
                          type="number"
                          value={row.lecturePlan || ''}
                          onChange={(e) => updateRow(row.id, { lecturePlan: Number(e.target.value) })}
                          className="w-full h-12 bg-transparent text-center font-bold text-yellow-800 focus:outline-none"
                        />
                      </td>
                      <td className="p-1 border-r border-slate-200 bg-green-100/40">
                        <div className="w-full h-12 flex items-center justify-center font-black text-green-700 text-sm">
                          {row.lectureTotal || 0}
                        </div>
                      </td>

                      {/* Lab - Plan(Yellow), Total(Green), Credits(Green) */}
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/50">
                        <input
                          type="number"
                          value={row.labPlan || ''}
                          onChange={(e) => updateRow(row.id, { labPlan: Number(e.target.value) })}
                          className="w-full h-12 bg-transparent text-center font-bold text-yellow-800 focus:outline-none"
                        />
                      </td>
                      <td className="p-1 border-r border-slate-200 bg-green-100/40 text-center font-black text-green-700">
                         {row.labTotal || 0}
                      </td>
                      <td className="p-1 border-r border-slate-200 bg-green-200/40 text-center font-black text-green-800">
                         {row.labCredits || 0}
                      </td>

                      {/* SROP - White */}
                      <td className="p-1 border-r border-slate-200 bg-white">
                         <input
                          type="number"
                          value={row.srop || ''}
                          onChange={(e) => updateRow(row.id, { srop: Number(e.target.value) })}
                          className="w-full h-12 text-center border-2 border-slate-50 rounded-md focus:border-indigo-300 focus:outline-none"
                        />
                      </td>

                      {/* Total Credits - Green */}
                      <td className="p-1 border-r border-slate-200 bg-green-400/20 font-black text-center text-indigo-900 border-2 border-green-200/50 shadow-inner">
                        <span className="text-sm">{row.totalCredits || 0}</span>
                      </td>

                      {/* Course Work - Yellow/Green */}
                      <td className="p-1 border-r border-slate-200 bg-yellow-50/40 text-center">
                        <div className="h-12 flex items-center justify-center">
                          <input
                            type="checkbox"
                            checked={row.courseWork}
                            onChange={(e) => updateRow(row.id, { courseWork: e.target.checked })}
                            className="w-6 h-6 accent-yellow-500 rounded-lg cursor-pointer transform scale-125"
                          />
                        </div>
                      </td>
                      <td className="p-1 border-r border-slate-200 bg-green-100/40 text-center font-bold text-green-800">
                         {row.courseWorkCredits || 0}
                      </td>

                      {/* PPS - Blue/Purple */}
                      <td className="p-2 min-w-[240px] bg-purple-50/20">
                        <div className="flex items-center gap-3">
                          <div className={cn(
                            "w-10 h-10 rounded-xl flex items-center justify-center text-white text-xs font-black shadow-md transition-all shrink-0",
                            row.ppsName ? "bg-gradient-to-br from-purple-500 to-indigo-600 rotate-3" : "bg-slate-300"
                          )}>
                             {row.ppsName ? row.ppsName[0].toUpperCase() : <Users size={16}/>}
                          </div>
                          <input
                            type="text"
                            placeholder="ФИО преподавателя"
                            value={row.ppsName}
                            onChange={(e) => updateRow(row.id, { ppsName: e.target.value })}
                            className="flex-1 bg-purple-50 border-2 border-purple-100 rounded-lg p-2.5 text-xs font-semibold focus:border-purple-400 focus:outline-none placeholder:text-purple-300/60"
                          />
                        </div>
                      </td>

                      <td className="p-1 text-center bg-white group-hover:bg-slate-50">
                        <button 
                          onClick={() => removeRow(row.id)}
                          className="p-3 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-full transition-all active:scale-90"
                        >
                          <Trash2 size={20} />
                        </button>
                      </td>
                    </motion.tr>
                  ))}
                </AnimatePresence>
              )}
            </tbody>
          </table>
        </div>
      </main>

      <footer className="max-w-[1920px] mx-auto mt-20 mb-32 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-8">
        {AUTHORITIES.map((auth, idx) => (
          <motion.div 
            key={idx} 
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="p-6 bg-white rounded-2xl border border-slate-100 shadow-sm hover:shadow-md transition-shadow flex flex-col gap-4 relative overflow-hidden"
          >
            <div className="absolute top-0 right-0 w-16 h-16 bg-slate-50 rounded-bl-3xl -mr-8 -mt-8"></div>
            <span className="text-[10px] font-black text-indigo-400 uppercase tracking-widest">{auth.role}</span>
            <div className="flex flex-col gap-3">
              <div className="h-px bg-gradient-to-r from-transparent via-slate-200 to-transparent"></div>
              <div className="flex items-center justify-between min-h-[40px]">
                <div className="w-1/2 border-b border-dashed border-slate-300 h-6"></div>
                <span className="text-xs font-bold text-slate-800 text-right">{auth.name}</span>
              </div>
            </div>
          </motion.div>
        ))}
      </footer>

      {/* Legend */}
      <div className="fixed bottom-6 right-6 flex flex-col gap-2 p-4 bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl z-50">
        <div className="text-[9px] font-bold text-slate-400 mb-2 uppercase tracking-widest border-b border-slate-800 pb-2">Условные обозначения</div>
        <div className="flex items-center gap-3 group">
          <div className="w-4 h-4 bg-yellow-200 rounded-sm shadow-[0_0_8px_rgba(254,240,138,0.5)] transition-transform group-hover:scale-110"></div>
          <span className="text-[10px] text-white/80 font-bold">Желтая: Из РУП (План)</span>
        </div>
        <div className="flex items-center gap-3 group">
          <div className="w-4 h-4 bg-white rounded-sm shadow-inner ring-1 ring-white/20 transition-transform group-hover:scale-110"></div>
          <span className="text-[10px] text-white/80 font-bold">Белая: Ввод вручную</span>
        </div>
        <div className="flex items-center gap-3 group">
          <div className="w-4 h-4 bg-green-400 rounded-sm shadow-[0_0_8px_rgba(74,222,128,0.5)] transition-transform group-hover:scale-110"></div>
          <span className="text-[10px] text-white/80 font-bold">Зеленая: Авт. Расчет</span>
        </div>
        <div className="flex items-center gap-3 group">
          <div className="w-4 h-4 bg-purple-500 rounded-sm shadow-[0_0_8px_rgba(168,85,247,0.5)] transition-transform group-hover:scale-110"></div>
          <span className="text-[10px] text-white/80 font-bold">Фиолет: ППС Штатный</span>
        </div>
      </div>
    </div>
  );
}

