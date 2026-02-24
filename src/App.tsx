import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { compareTwoStrings } from 'string-similarity';
import { UploadCloud, FileSpreadsheet, AlertCircle, CheckCircle, Download, RefreshCw, FileText, Info } from 'lucide-react';

const safeCompare = (str1: string, str2: string) => {
  if (!str1 || !str2) return 0;
  if (str1 === str2) return 1;
  if (str1.length < 2 || str2.length < 2) {
    return str1 === str2 ? 1 : 0;
  }
  return compareTwoStrings(str1, str2);
};

interface RowData {
  'Id enroll'?: string | number;
  "Code d'employé"?: string | number;
  'Nom, prénom'?: string;
  'BioStar'?: string;
  'BioStar_corrected'?: string;
  [key: string]: any;
}

interface CorrectionRecord {
  rowNumber: number;
  id: string | number;
  nom: string;
  originalBioStar: string;
  correctedBioStar: string;
  similarity: number;
}

interface UnmatchedRecord {
  rowNumber: number;
  id: string | number;
  nom: string;
  bioStar: string;
  similarity: number;
}

export default function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedData, setProcessedData] = useState<RowData[] | null>(null);
  const [corrections, setCorrections] = useState<CorrectionRecord[]>([]);
  const [unmatched, setUnmatched] = useState<UnmatchedRecord[]>([]);
  const [stats, setStats] = useState({ total: 0, perfect: 0, corrected: 0, unmatched: 0 });
  const [activeTab, setActiveTab] = useState<'corrections' | 'unmatched'>('corrections');

  const processFile = (file: File) => {
    setIsProcessing(true);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const jsonData = XLSX.utils.sheet_to_json<RowData>(worksheet, { defval: '' });
        
        const newCorrections: CorrectionRecord[] = [];
        const newUnmatched: UnmatchedRecord[] = [];
        let perfectCount = 0;
        
        // Pre-compute all unique valid names for global search
        const allNomsOriginal = Array.from(new Set(jsonData.map(r => (r['Nom, prénom'] || '').toString().trim()).filter(Boolean)));
        const allNomsLower = allNomsOriginal.map(n => n.toLowerCase());
        
        const processed = jsonData.map((row, index) => {
          const newRow = { ...row };
          const nom = (row['Nom, prénom'] || '').toString().trim();
          const bioStar = (row['BioStar'] || '').toString().trim();
          const rowNum = index + 2; // +1 for 0-index, +1 for header
          const id = row['Id enroll'] || row["Code d'employé"] || `Row ${rowNum}`;
          
          if (!nom && !bioStar) {
            newRow['BioStar_corrected'] = '';
            return newRow;
          }
          
          if (nom.toLowerCase() === bioStar.toLowerCase()) {
            newRow['BioStar_corrected'] = nom;
            perfectCount++;
            return newRow;
          }
          
          let similarity = safeCompare(nom.toLowerCase(), bioStar.toLowerCase());
          let bestMatchNom = nom;
          
          // If row-by-row match is not reliable, search globally across all names
          if (similarity < 0.85 && bioStar && allNomsLower.length > 0) {
            let bestGlobalScore = 0;
            let bestGlobalIndex = -1;
            
            for (let i = 0; i < allNomsLower.length; i++) {
              const score = safeCompare(bioStar.toLowerCase(), allNomsLower[i]);
              if (score > bestGlobalScore) {
                bestGlobalScore = score;
                bestGlobalIndex = i;
              }
            }
            
            if (bestGlobalScore > similarity) {
              similarity = bestGlobalScore;
              bestMatchNom = allNomsOriginal[bestGlobalIndex];
            }
          }
          
          if (similarity >= 0.85) {
            newRow['BioStar_corrected'] = bestMatchNom;
            newCorrections.push({
              rowNumber: rowNum,
              id,
              nom: bestMatchNom,
              originalBioStar: bioStar,
              correctedBioStar: bestMatchNom,
              similarity
            });
          } else {
            newRow['BioStar_corrected'] = bioStar;
            newUnmatched.push({
              rowNumber: rowNum,
              id,
              nom: bestMatchNom,
              bioStar,
              similarity
            });
          }
          
          return newRow;
        });
        
        setProcessedData(processed);
        setCorrections(newCorrections);
        setUnmatched(newUnmatched);
        setStats({
          total: processed.length,
          perfect: perfectCount,
          corrected: newCorrections.length,
          unmatched: newUnmatched.length
        });
      } catch (error) {
        console.error("Error processing file:", error);
        alert("An error occurred while processing the file. Please ensure it's a valid Excel file with the required columns.");
      } finally {
        setIsProcessing(false);
      }
    };
    
    reader.readAsArrayBuffer(file);
  };

  const onDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        processFile(file);
      } else {
        alert("Please upload a valid Excel file (.xlsx or .xls)");
      }
    }
  }, []);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  const downloadExcel = () => {
    if (!processedData) return;
    
    const worksheet = XLSX.utils.json_to_sheet(processedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Corrected Data");
    
    XLSX.writeFile(workbook, "CER_corrected.xlsx");
  };

  const downloadReport = () => {
    if (!processedData) return;
    
    let reportContent = `DataReconcile Summary Report\n`;
    reportContent += `==============================\n\n`;
    reportContent += `Total Rows Processed: ${stats.total}\n`;
    reportContent += `Perfect Matches: ${stats.perfect}\n`;
    reportContent += `Auto-Corrected: ${stats.corrected}\n`;
    reportContent += `Unmatched (Flagged): ${stats.unmatched}\n\n`;
    
    reportContent += `--- CORRECTIONS MADE (${stats.corrected}) ---\n`;
    if (corrections.length === 0) {
      reportContent += `No corrections were necessary.\n`;
    } else {
      corrections.forEach(c => {
        reportContent += `Row ${c.rowNumber} | ID: ${c.id} | Truth: "${c.nom}" | Original: "${c.originalBioStar}" | Similarity: ${(c.similarity * 100).toFixed(1)}%\n`;
      });
    }
    
    reportContent += `\n--- UNMATCHED ENTRIES (${stats.unmatched}) ---\n`;
    if (unmatched.length === 0) {
      reportContent += `All entries were successfully matched!\n`;
    } else {
      unmatched.forEach(u => {
        reportContent += `Row ${u.rowNumber} | ID: ${u.id} | Truth: "${u.nom}" | BioStar: "${u.bioStar}" | Best Match Similarity: ${(u.similarity * 100).toFixed(1)}%\n`;
      });
    }
    
    const blob = new Blob([reportContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'CER_summary_report.txt';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const downloadUnmatchedExcel = () => {
    if (unmatched.length === 0) return;
    
    const exportData = unmatched.map(u => ({
      'Row': u.rowNumber,
      'ID': u.id,
      'Nom, prénom (Truth)': u.nom,
      'BioStar': u.bioStar,
      'Similarity': `${(u.similarity * 100).toFixed(1)}%`
    }));
    
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Unmatched Entries");
    
    XLSX.writeFile(workbook, "CER_unmatched.xlsx");
  };

  const resetApp = () => {
    setProcessedData(null);
    setCorrections([]);
    setUnmatched([]);
    setStats({ total: 0, perfect: 0, corrected: 0, unmatched: 0 });
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-indigo-100 selection:text-indigo-900">
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2 text-indigo-600">
            <FileSpreadsheet className="w-6 h-6" />
            <h1 className="text-xl font-semibold tracking-tight text-slate-900">DataReconcile</h1>
          </div>
          {processedData && (
            <button 
              onClick={resetApp}
              className="flex items-center gap-2 text-sm font-medium text-slate-600 hover:text-slate-900 transition-colors"
            >
              <RefreshCw className="w-4 h-4" />
              Process New File
            </button>
          )}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8">
        {!processedData ? (
          <div className="max-w-2xl mx-auto mt-12">
            <div className="text-center mb-8">
              <h2 className="text-3xl font-semibold tracking-tight mb-3">Upload your Excel file</h2>
              <p className="text-slate-500">
                We'll compare the <span className="font-medium text-slate-700">"Nom, prénom"</span> and <span className="font-medium text-slate-700">"BioStar"</span> columns, fix minor spelling differences, and generate a clean report.
              </p>
            </div>

            <div 
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              className={`relative border-2 border-dashed rounded-2xl p-12 text-center transition-all duration-200 ease-in-out ${
                isDragging 
                  ? 'border-indigo-500 bg-indigo-50/50 scale-[1.02]' 
                  : 'border-slate-300 hover:border-slate-400 bg-white hover:bg-slate-50'
              }`}
            >
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                onChange={handleFileInput}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <div className="flex flex-col items-center gap-4 pointer-events-none">
                <div className={`p-4 rounded-full ${isDragging ? 'bg-indigo-100 text-indigo-600' : 'bg-slate-100 text-slate-500'}`}>
                  <UploadCloud className="w-8 h-8" />
                </div>
                <div>
                  <p className="text-lg font-medium text-slate-900">
                    {isDragging ? 'Drop file here' : 'Click or drag file to upload'}
                  </p>
                  <p className="text-sm text-slate-500 mt-1">Supports .xlsx and .xls</p>
                </div>
              </div>
            </div>

            <div className="mt-8 bg-blue-50 border border-blue-100 rounded-xl p-4 flex gap-3 text-blue-800">
              <Info className="w-5 h-5 shrink-0 mt-0.5 text-blue-600" />
              <div className="text-sm leading-relaxed">
                <p className="font-medium mb-1">Required Columns:</p>
                <ul className="list-disc list-inside space-y-1 text-blue-700/80">
                  <li>Id enroll (or Code d'employé)</li>
                  <li>Nom, prénom</li>
                  <li>BioStar</li>
                </ul>
              </div>
            </div>
          </div>
        ) : (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-2xl font-semibold tracking-tight">Reconciliation Report</h2>
                <p className="text-slate-500 mt-1">Processed {stats.total} rows successfully.</p>
              </div>
              <div className="flex items-center gap-3">
                <button 
                  onClick={downloadReport}
                  className="inline-flex items-center justify-center gap-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 px-5 py-2.5 rounded-xl font-medium transition-colors shadow-sm"
                >
                  <FileText className="w-4 h-4" />
                  Download Report
                </button>
                <button 
                  onClick={downloadExcel}
                  className="inline-flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-xl font-medium transition-colors shadow-sm shadow-indigo-600/20"
                >
                  <Download className="w-4 h-4" />
                  Download Corrected Excel
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                <div className="text-slate-500 text-sm font-medium mb-1">Total Rows</div>
                <div className="text-3xl font-semibold">{stats.total}</div>
              </div>
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                <div className="text-emerald-600 text-sm font-medium mb-1 flex items-center gap-1.5">
                  <CheckCircle className="w-4 h-4" /> Perfect Matches
                </div>
                <div className="text-3xl font-semibold text-emerald-700">{stats.perfect}</div>
              </div>
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                <div className="text-amber-600 text-sm font-medium mb-1 flex items-center gap-1.5">
                  <RefreshCw className="w-4 h-4" /> Auto-Corrected
                </div>
                <div className="text-3xl font-semibold text-amber-700">{stats.corrected}</div>
              </div>
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                <div className="text-rose-600 text-sm font-medium mb-1 flex items-center gap-1.5">
                  <AlertCircle className="w-4 h-4" /> Unmatched (Flagged)
                </div>
                <div className="text-3xl font-semibold text-rose-700">{stats.unmatched}</div>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[600px]">
              <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50/50 px-4 pt-4">
                <div className="flex gap-4">
                  <button
                    onClick={() => setActiveTab('corrections')}
                    className={`pb-3 px-2 text-sm font-medium border-b-2 transition-colors ${
                      activeTab === 'corrections' 
                        ? 'border-indigo-600 text-indigo-600' 
                        : 'border-transparent text-slate-500 hover:text-slate-700'
                    }`}
                  >
                    Corrections Made ({stats.corrected})
                  </button>
                  <button
                    onClick={() => setActiveTab('unmatched')}
                    className={`pb-3 px-2 text-sm font-medium border-b-2 transition-colors ${
                      activeTab === 'unmatched' 
                        ? 'border-rose-500 text-rose-600' 
                        : 'border-transparent text-slate-500 hover:text-slate-700'
                    }`}
                  >
                    Unmatched Entries ({stats.unmatched})
                  </button>
                </div>
                {activeTab === 'unmatched' && unmatched.length > 0 && (
                  <button
                    onClick={downloadUnmatchedExcel}
                    className="mb-2 inline-flex items-center gap-1.5 text-xs font-medium text-rose-600 hover:text-rose-700 bg-rose-50 hover:bg-rose-100 px-3 py-1.5 rounded-lg transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Export Unmatched
                  </button>
                )}
              </div>

              <div className="flex-1 overflow-auto p-0">
                {activeTab === 'corrections' ? (
                  corrections.length > 0 ? (
                    <table className="w-full text-sm text-left">
                      <thead className="text-xs text-slate-500 uppercase bg-slate-50 sticky top-0">
                        <tr>
                          <th className="px-6 py-3 font-medium">Row</th>
                          <th className="px-6 py-3 font-medium">ID</th>
                          <th className="px-6 py-3 font-medium">Nom, prénom (Truth)</th>
                          <th className="px-6 py-3 font-medium">Original BioStar</th>
                          <th className="px-6 py-3 font-medium">Similarity</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {corrections.map((c, i) => (
                          <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-6 py-3 text-slate-500">{c.rowNumber}</td>
                            <td className="px-6 py-3 font-mono text-xs text-slate-500">{c.id}</td>
                            <td className="px-6 py-3 font-medium text-slate-900">{c.nom}</td>
                            <td className="px-6 py-3 text-amber-600 line-through decoration-amber-300">{c.originalBioStar}</td>
                            <td className="px-6 py-3">
                              <span className="inline-flex items-center px-2 py-1 rounded-md bg-emerald-50 text-emerald-700 text-xs font-medium">
                                {(c.similarity * 100).toFixed(1)}%
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <div className="h-full flex flex-col items-center justify-center text-slate-500 p-8">
                      <CheckCircle className="w-12 h-12 text-slate-200 mb-3" />
                      <p>No corrections were necessary.</p>
                    </div>
                  )
                ) : (
                  unmatched.length > 0 ? (
                    <table className="w-full text-sm text-left">
                      <thead className="text-xs text-slate-500 uppercase bg-slate-50 sticky top-0">
                        <tr>
                          <th className="px-6 py-3 font-medium">Row</th>
                          <th className="px-6 py-3 font-medium">ID</th>
                          <th className="px-6 py-3 font-medium">Nom, prénom (Truth)</th>
                          <th className="px-6 py-3 font-medium">BioStar</th>
                          <th className="px-6 py-3 font-medium">Similarity</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {unmatched.map((u, i) => (
                          <tr key={i} className="hover:bg-rose-50/30 transition-colors">
                            <td className="px-6 py-3 text-slate-500">{u.rowNumber}</td>
                            <td className="px-6 py-3 font-mono text-xs text-slate-500">{u.id}</td>
                            <td className="px-6 py-3 font-medium text-slate-900">{u.nom || <span className="text-slate-400 italic">Empty</span>}</td>
                            <td className="px-6 py-3 text-rose-600 font-medium">{u.bioStar || <span className="text-slate-400 italic">Empty</span>}</td>
                            <td className="px-6 py-3">
                              <span className="inline-flex items-center px-2 py-1 rounded-md bg-rose-50 text-rose-700 text-xs font-medium">
                                {(u.similarity * 100).toFixed(1)}%
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <div className="h-full flex flex-col items-center justify-center text-slate-500 p-8">
                      <CheckCircle className="w-12 h-12 text-slate-200 mb-3" />
                      <p>All entries were successfully matched!</p>
                    </div>
                  )
                )}
              </div>
            </div>
          </div>
        )}
      </main>
      
      {isProcessing && (
        <div className="fixed inset-0 bg-white/80 backdrop-blur-sm z-50 flex items-center justify-center">
          <div className="flex flex-col items-center gap-4">
            <div className="w-12 h-12 border-4 border-indigo-200 border-t-indigo-600 rounded-full animate-spin"></div>
            <p className="text-lg font-medium text-slate-900 animate-pulse">Processing Excel file...</p>
          </div>
        </div>
      )}
    </div>
  );
}
