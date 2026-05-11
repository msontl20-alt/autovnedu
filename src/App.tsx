import {useState, useCallback, useMemo, useRef, useEffect} from 'react';
import * as XLSX from 'xlsx';
import {motion, AnimatePresence} from 'motion/react';
import {
  Upload,
  FileSpreadsheet,
  Trash2,
  Wand2,
  Download,
  Copy,
  GraduationCap,
  CheckCircle,
  X,
  Book,
  Users,
  Calendar,
  PlusCircle,
  FileCheck,
  RefreshCw,
  RotateCcw,
  Save
} from 'lucide-react';
import {cn} from '@/src/lib/utils';
import {
  ALL_SUBJECTS_BANK,
  competenciesGood,
  competenciesFair,
  competenciesPoor,
  qualitiesGood,
  qualitiesFair,
  qualitiesPoor
} from './constants';

interface StudentData {
  _rowIndex: number;
  [key: string]: any;
}

export default function App() {
  const [rawDataAOA, setRawDataAOA] = useState<any[][]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [headerRowIndex, setHeaderRowIndex] = useState(-1);
  const [studentsData, setStudentsData] = useState<StudentData[]>([]);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [originalFileName, setOriginalFileName] = useState("");
  const [isProcessed, setIsProcessed] = useState(false);
  const [selectedPeriods, setSelectedPeriods] = useState<Set<"GK1" | "CK1" | "GK2" | "CK2">>(new Set(["CK1"]));
  const [filterEmptyComments, setFilterEmptyComments] = useState(false);
  const [showConfirm, setShowConfirm] = useState<{type: 'all' | 'comments', action: () => void} | null>(null);

  // Config State
  const [subject, setSubject] = useState("Toán");
  const [gradeLevel, setGradeLevel] = useState("general");
  const [colMapping, setColMapping] = useState({
    name: "",
    studentId: "",
    gender: "",
    teacher: "",
    periods: {
      GK1: { score: "", level: "", comment: "" },
      CK1: { score: "", level: "", comment: "" },
      GK2: { score: "", level: "", comment: "" },
      CK2: { score: "", level: "", comment: "" },
    }
  });
  const [autoLevel, setAutoLevel] = useState(true);
  const [addQualities, setAddQualities] = useState(true);

  // UI State
  const [toast, setToast] = useState<{msg: string, type: 'success' | 'error'} | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const showToast = (msg: string, type: 'success' | 'error' = 'success') => {
    setToast({msg, type});
    setTimeout(() => setToast(null), 3000);
  };

  useEffect(() => {
    const saved = localStorage.getItem('report_helper_default_config');
    if (saved) {
      try {
        const config = JSON.parse(saved);
        if (config.subject) setSubject(config.subject);
        if (config.gradeLevel) setGradeLevel(config.gradeLevel);
      } catch (e) {
        console.error("Error loading initial config", e);
      }
    }
  }, []);

  const removeAccents = (str: string) => {
    return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D');
  };

  const handleFiles = (file: File) => {
    setOriginalFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, {type: 'array', cellStyles: true});
        const sn = wb.SheetNames[0];
        const ws = wb.Sheets[sn];
        const aoa = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ""}) as any[][];

        setWorkbook(wb);
        setSheetName(sn);
        setRawDataAOA(aoa);

        // Find Headers
        let hIdx = -1;
        for (let i = 0; i < Math.min(15, aoa.length); i++) {
          const row = aoa[i];
          if (row.some(cell => typeof cell === 'string' && (cell.toLowerCase().includes('tên') || cell.toLowerCase().includes('họ và')))) {
            hIdx = i;
            break;
          }
        }
        if (hIdx === -1) hIdx = 0;
        setHeaderRowIndex(hIdx);

        const h = aoa[hIdx].map((val, i) => val ? String(val).trim() : `Column_${i}`);
        setHeaders(h);

        // Auto detect config from file name and headers
        const cleanName = removeAccents(file.name.toLowerCase());
        
        // 1. Subject Detection (Filename + Content scan)
        let detectedSubject = "Toán";
        const subjectKeywords: Record<string, string[]> = {
          "Toán": ["toan", "math"],
          "Tiếng Việt": ["tieng viet", " tv ", "mon tv"],
          "Tin học": ["tin hoc", " tin ", "it"],
          "Tiếng Anh": ["tieng anh", " anh ", "english", " t.anh"],
          "Đạo đức": ["dao duc", " dd "],
          "Tự nhiên và Xã hội": ["tu nhien", "tnxh", "tn&xh"],
          "Khoa học": ["khoa hoc", " kh "],
          "Lịch sử và Địa lí": ["lich su", "dia li", "lsdl", "ls&dl", "ls dl"],
          "Công nghệ": ["cong nghe", " cn "],
          "Hoạt động trải nghiệm": ["hoat dong", "hdtn", "trai nghiem"],
          "Âm nhạc": ["am nhac", " an "],
          "Mĩ thuật": ["mi thuat", " mt "],
          "Giáo dục thể chất": ["the chat", "gdtc", "the duc"],
        };

        const detectFromStr = (str: string) => {
          for (const [sub, keywords] of Object.entries(subjectKeywords)) {
            if (keywords.some(k => str.includes(k))) return sub;
          }
          return null;
        };

        const fromName = detectFromStr(cleanName);
        if (fromName) {
          detectedSubject = fromName;
        } else {
          // Scan content for "Môn: Toán" or similar
          for (let i = 0; i < Math.min(15, aoa.length); i++) {
            const rowStr = removeAccents(aoa[i].join(' ').toLowerCase());
            const fromContent = detectFromStr(rowStr);
            if (fromContent) {
              detectedSubject = fromContent;
              break;
            }
          }
        }
        setSubject(detectedSubject);

        // 2. Grade Detection (Filename first, then deep scan)
        let detectedGrade = "general";
        const gradeMatch = cleanName.match(/(lop|khoi|k)\s*([12345])/);
        if (gradeMatch) {
          detectedGrade = gradeMatch[2];
        } else {
          // Deep scan first 10 rows for "Lớp: 3A" or "Khối 2"
          for (let i = 0; i < Math.min(10, aoa.length); i++) {
            const rowStr = removeAccents(aoa[i].join(' ').toLowerCase());
            const m = rowStr.match(/(lop|khoi)\s*([12345])/);
            if (m) {
              detectedGrade = m[2];
              break;
            }
          }
        }
        // Priority: Auto Detection first, but we can potentially merge with saved if headers match
        const mapping = detectMappings(h);
        
        // Merge with defaults if they exist and are valid for this file
        const savedStr = localStorage.getItem('report_helper_default_config');
        if (savedStr) {
          try {
            const saved = JSON.parse(savedStr);
            const hSet = new Set(h);
            
            // Apply saved mapping for fields not strongly detected or always prefer saved if header matches
            if (saved.colMapping) {
              const sm = saved.colMapping;
              if (sm.name && hSet.has(sm.name)) mapping.name = sm.name;
              if (sm.studentId && hSet.has(sm.studentId)) mapping.studentId = sm.studentId;
              if (sm.gender && hSet.has(sm.gender)) mapping.gender = sm.gender;
              if (sm.teacher && hSet.has(sm.teacher)) mapping.teacher = sm.teacher;
              
              (["GK1", "CK1", "GK2", "CK2"] as const).forEach(p => {
                if (sm.periods[p].score && hSet.has(sm.periods[p].score)) mapping.periods[p].score = sm.periods[p].score;
                if (sm.periods[p].level && hSet.has(sm.periods[p].level)) mapping.periods[p].level = sm.periods[p].level;
                if (sm.periods[p].comment && hSet.has(sm.periods[p].comment)) mapping.periods[p].comment = sm.periods[p].comment;
              });
            }
            
            // Only use saved subject/grade if detection was "general/Toán" (weak detection)
            if (detectedSubject === "Toán" && saved.subject) detectedSubject = saved.subject;
            if (detectedGrade === "general" && saved.gradeLevel) detectedGrade = saved.gradeLevel;
            
          } catch (e) {
            console.error("Error loading default config during upload", e);
          }
        }

        setGradeLevel(detectedGrade);
        setSubject(detectedSubject);
        setColMapping(mapping);

        // 3. Period Detection
        const hStr = h.join(' ').toLowerCase();
        const activePeriods = new Set<"GK1" | "CK1" | "GK2" | "CK2">();
        
        if (cleanName.includes('gk1') || hStr.includes('gk1')) activePeriods.add("GK1");
        if (cleanName.includes('ck1') || cleanName.includes('hoc ky 1') || hStr.includes('ck1')) activePeriods.add("CK1");
        if (cleanName.includes('gk2') || hStr.includes('gk2')) activePeriods.add("GK2");
        if (cleanName.includes('ck2') || cleanName.includes('hoc ky 2') || hStr.includes('ck2')) activePeriods.add("CK2");
        
        if (activePeriods.size === 0) activePeriods.add("CK1");
        setSelectedPeriods(activePeriods);

        // Extract student data
        const students: StudentData[] = [];
        for (let i = hIdx + 1; i < aoa.length; i++) {
          const row = aoa[i];
          if (row.some(cell => cell !== "")) {
            let s: StudentData = {_rowIndex: i};
            h.forEach((col, colIndex) => {
              s[col] = row[colIndex] ?? "";
            });
            students.push(s);
          }
        }
        setStudentsData(students);
        showToast("Tải file thành công!");
      } catch (err) {
        console.error(err);
        showToast("Lỗi đọc file Excel!", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Config Actions
  const detectMappings = (h: string[]) => {
    const mapping = {
      name: "", 
      studentId: "", 
      gender: "", 
      teacher: "",
      periods: {
        GK1: { score: "", level: "", comment: "" },
        CK1: { score: "", level: "", comment: "" },
        GK2: { score: "", level: "", comment: "" },
        CK2: { score: "", level: "", comment: "" },
      }
    };

    let lastDetectedPeriod: "GK1" | "CK1" | "GK2" | "CK2" | null = null;

    h.forEach((col) => {
      const raw = col.toLowerCase().trim();
      const clean = removeAccents(raw);
      
      // 1. Detect Period from name
      const pMatch = clean.match(/(gk1|ck1|gk2|ck2)/);
      const periodInName = pMatch ? pMatch[0].toUpperCase() as "GK1"|"CK1"|"GK2"|"CK2" : null;
      if (periodInName) lastDetectedPeriod = periodInName;

      // 2. Global Fields
      if (clean.includes('ho va ten') || clean.includes('ho ten') || clean === 'ten' || clean.includes('ten hoc sinh') || clean === 'hoten' || clean === 'tenhs') {
        if (!mapping.name) mapping.name = col;
      }
      else if (clean.includes('ma hoc sinh') || clean.includes('ma so') || clean === 'mshs' || clean === 'id' || clean === 'ma hs' || clean.includes('ma dinh danh') || clean === 'uid') {
        if (!mapping.studentId) mapping.studentId = col;
      }
      else if (clean.includes('gioi tinh') || clean === 'nu' || clean === 'nam/nu' || clean === 'nam' || clean === 'phai' || clean === 'gt' || clean.includes('phai / gt') || clean.includes('gioi')) {
        if (!mapping.gender) mapping.gender = col;
      }
      else if (clean.includes('giao vien chu nhiem') || clean === 'gvcn' || clean.includes('ten giao vien') || clean === 'giao vien' || clean === 'gv' || clean.includes('nguoi cham') || clean.includes('thay/co') || clean.includes('gv giang day') || clean.includes('giao vien mon')) {
        if (!mapping.teacher) mapping.teacher = col;
      }
      
      // 3. Period Specific Fields
      const targetPeriod = periodInName || lastDetectedPeriod || "CK1";
      
      // Comment detection (Prioritized because NX columns sometimes contain other keywords)
      const isComment = clean.includes('nhan xet') || clean.includes('nx') || clean.includes('danh gia') || clean.includes('loi phe') || clean.includes('nhanxet') || clean.includes('ghi chu') || clean.includes('noi dung');
      
      // Level detection
      const isLevel = !isComment && (clean.includes('muc') || clean.includes('xl') || clean.includes('xep loai') || clean === 'muc dat' || (clean.startsWith('xl') && (clean.includes('k1') || clean.includes('k2') || clean.includes('ck'))));
      
      // Score detection
      const isScore = !isComment && !isLevel && (clean.includes('diem') || clean.includes('ktdk') || clean.includes('kt ') || clean.includes('kiem tra') || (clean.startsWith('kt') && (clean.includes('k1') || clean.includes('k2') || clean.includes('ck'))));

      if (isComment) {
        if (!mapping.periods[targetPeriod].comment || periodInName) mapping.periods[targetPeriod].comment = col;
      } else if (isLevel) {
        if (!mapping.periods[targetPeriod].level || periodInName) mapping.periods[targetPeriod].level = col;
      } else if (isScore) {
        if (!mapping.periods[targetPeriod].score || periodInName) mapping.periods[targetPeriod].score = col;
      }
    });

    return mapping;
  };

  const autoMap = () => {
    if (!headers.length) return;
    setColMapping(detectMappings(headers));
    showToast("Đã tự động ánh xạ lại các cột!");
  };

  const clearMappings = () => {
    setColMapping({
      name: "",
      studentId: "",
      gender: "",
      teacher: "",
      periods: {
        GK1: { score: "", level: "", comment: "" },
        CK1: { score: "", level: "", comment: "" },
        GK2: { score: "", level: "", comment: "" },
        CK2: { score: "", level: "", comment: "" },
      }
    });
    showToast("Đã xóa toàn bộ ánh xạ!");
  };

  const saveAsDefault = () => {
    localStorage.setItem('report_helper_default_config', JSON.stringify({
      subject,
      gradeLevel,
      colMapping
    }));
    showToast("Đã lưu cấu hình làm mặc định!");
  };

  const loadDefault = () => {
    const saved = localStorage.getItem('report_helper_default_config');
    if (saved) {
      try {
        const config = JSON.parse(saved);
        if (config.subject) setSubject(config.subject);
        if (config.gradeLevel) setGradeLevel(config.gradeLevel);
        if (config.colMapping) {
          // If we have headers, we should only apply mappings that exist in the current headers
          if (headers.length > 0) {
            const mapped = { ...config.colMapping };
            const hSet = new Set(headers);
            if (mapped.name && !hSet.has(mapped.name)) mapped.name = "";
            if (mapped.studentId && !hSet.has(mapped.studentId)) mapped.studentId = "";
            if (mapped.gender && !hSet.has(mapped.gender)) mapped.gender = "";
            if (mapped.teacher && !hSet.has(mapped.teacher)) mapped.teacher = "";
            
            (["GK1", "CK1", "GK2", "CK2"] as const).forEach(p => {
              if (mapped.periods[p].score && !hSet.has(mapped.periods[p].score)) mapped.periods[p].score = "";
              if (mapped.periods[p].level && !hSet.has(mapped.periods[p].level)) mapped.periods[p].level = "";
              if (mapped.periods[p].comment && !hSet.has(mapped.periods[p].comment)) mapped.periods[p].comment = "";
            });
            setColMapping(mapped);
          } else {
            setColMapping(config.colMapping);
          }
        }
        showToast("Đã tải cấu hình mặc định!");
      } catch (e) {
        showToast("Lỗi khi tải mặc định!", "error");
      }
    } else {
      showToast("Chưa có cấu hình mặc định nào được lưu!", "error");
    }
  };

  const processComments = () => {
    if (!colMapping.name) {
      showToast("Chưa cấu hình cột Họ tên!", "error");
      return;
    }

    if (selectedPeriods.size === 0) {
      showToast("Vui lòng chọn ít nhất một kỳ đánh giá!", "error");
      return;
    }

    const subjectData = ALL_SUBJECTS_BANK[subject] || ALL_SUBJECTS_BANK["Toán"];
    const bank = subjectData[gradeLevel] || subjectData["general"];
    const getRandom = (arr: string[]) => arr[Math.floor(Math.random() * arr.length)];

    const newStudents = studentsData.map(s => {
      const name = String(s[colMapping.name] || "").trim();
      if (!name || name.toLowerCase().includes('tổng')) return s;

      let updatedStudent = { ...s };

      selectedPeriods.forEach(period => {
        const pMap = colMapping.periods[period];
        if (!pMap.comment) return;

        let scoreStr = pMap.score ? String(s[pMap.score]).replace(',', '.') : "";
        let levelStr = pMap.level ? String(s[pMap.level]).toUpperCase().trim() : "";
        
        const isMidTerm = period.startsWith("GK");
        let category = "";
        if (scoreStr && !isNaN(parseFloat(scoreStr)) && !isMidTerm) {
          const score = parseFloat(scoreStr);
          if (score >= 9) category = "excellent";
          else if (score >= 7) category = "good";
          else if (score >= 5) category = "fair";
          else category = "poor";
        } else if (levelStr) {
          if (levelStr === 'T' || levelStr.includes('TỐT') || levelStr === 'HTX') category = "excellent";
          else if (levelStr === 'H' || levelStr.includes('HOÀN THÀNH')) category = Math.random() > 0.5 ? "good" : "fair";
          else if (levelStr === 'C' || levelStr.includes('CHƯA')) category = "poor";
        } else {
          // If no score/level, default to fair/good
          category = Math.random() > 0.3 ? "good" : "fair";
        }

        let comment = getRandom(bank[category]);
        if (addQualities) {
          const qualCat = (category === "excellent" || category === "good") ? "good" : (category === "fair" ? "fair" : "poor");
          const comp = getRandom(qualCat === "good" ? competenciesGood : (qualCat === "fair" ? competenciesFair : competenciesPoor));
          const qual = getRandom(qualCat === "good" ? qualitiesGood : (qualCat === "fair" ? qualitiesFair : qualitiesPoor));
          comment += ` ${comp} ${qual}`;
        }
        
        updatedStudent[pMap.comment] = comment;
        if (autoLevel && pMap.level && !isMidTerm && scoreStr) {
          const score = parseFloat(scoreStr);
          updatedStudent[pMap.level] = score >= 7 ? 'T' : (score >= 5 ? 'H' : 'C');
        }
      });

      return updatedStudent;
    });

    setStudentsData(newStudents);
    setIsProcessed(true);
    showToast(`Đã tạo nhận xét cho ${newStudents.length} học sinh!`);
  };

  const exportExcel = () => {
    if (!workbook || !sheetName) return;
    const executeExport = () => {
      const newData = [...rawDataAOA];
      
      studentsData.forEach(s => {
        const rIdx = s._rowIndex;
        
        // If filtering empty, check if we should skip updates for this row
        if (filterEmptyComments) {
          const hasAnyComment = Array.from(selectedPeriods).some(period => {
            const pMap = colMapping.periods[period];
            return (s[pMap.comment] || "").trim().length > 0;
          });
          if (!hasAnyComment) return;
        }

        selectedPeriods.forEach(period => {
          const pMap = colMapping.periods[period];
          const commentIdx = headers.indexOf(pMap.comment);
          const levelIdx = headers.indexOf(pMap.level);

          if (commentIdx !== -1) {
            while (newData[rIdx].length <= commentIdx) newData[rIdx].push("");
            newData[rIdx][commentIdx] = s[pMap.comment];
          }

          if (levelIdx !== -1 && autoLevel && !period.startsWith("GK")) {
            while (newData[rIdx].length <= levelIdx) newData[rIdx].push("");
            newData[rIdx][levelIdx] = s[pMap.level];
          }
        });
      });

      const ws = workbook.Sheets[sheetName];
      XLSX.utils.sheet_add_aoa(ws, newData, {origin: "A1"});
      
      const lastDot = originalFileName.lastIndexOf('.');
      const fileName = lastDot !== -1 
        ? `${originalFileName.substring(0, lastDot)} - xong${originalFileName.substring(lastDot)}`
        : `${originalFileName} - xong.xlsx`;

      XLSX.writeFile(workbook, fileName);
      showToast("Đã tải file Excel!");
      setShowConfirm(null);
    };

    setShowConfirm({type: 'all', action: executeExport});
  };

  const exportCommentsOnly = () => {
    if (!studentsData.length || selectedPeriods.size === 0) return;

    const executeExport = () => {
      const periodList = Array.from(selectedPeriods);
      const exportHeaders = ["STT", "Họ và Tên"];
      periodList.forEach(p => {
        exportHeaders.push(`${p} - Điểm`);
        exportHeaders.push(`${p} - Nhận xét`);
      });

      const exportRows = studentsData
        .filter(s => {
          const name = String(s[colMapping.name] || "");
          if (!name || name.toLowerCase().includes('tổng')) return false;
          if (filterEmptyComments) {
            return periodList.some(p => {
              const pMap = colMapping.periods[p];
              return (s[pMap.comment] || "").trim().length > 0;
            });
          }
          return true;
        })
        .map((s, idx) => {
          const row = [idx + 1, s[colMapping.name]];
          periodList.forEach(p => {
            const pMap = colMapping.periods[p];
            row.push(s[pMap.score] || "");
            row.push(s[pMap.comment] || "");
          });
          return row;
        });
      
      const aoa = [exportHeaders, ...exportRows];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Nhận xét");
      
      const fileName = `Danh sách nhận xét tổng hợp - ${subject}.xlsx`;
      XLSX.writeFile(wb, fileName);
      showToast("Đã xuất danh sách nhận xét riêng!");
      setShowConfirm(null);
    };

    setShowConfirm({type: 'comments', action: executeExport});
  };

  const copyAll = () => {
    if (selectedPeriods.size === 0) return;
    const period = Array.from(selectedPeriods)[0]; // Copy for the first selected period
    const pMap = colMapping.periods[period];
    
    const text = studentsData
      .filter(s => {
        const name = String(s[colMapping.name] || "");
        if (!name || name.toLowerCase().includes('tổng')) return false;
        if (filterEmptyComments) {
          return (s[pMap.comment] || "").trim().length > 0;
        }
        return true;
      })
      .map(s => `${s[colMapping.name]}:\t${s[pMap.comment]}`)
      .join('\n');
    navigator.clipboard.writeText(text);
    showToast(`Đã copy nhận xét của ${period}!`);
  };

  const reset = () => {
    setRawDataAOA([]);
    setStudentsData([]);
    setOriginalFileName("");
    setIsProcessed(false);
  };

  return (
    <div className="flex flex-col min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-blue-600 text-white shadow-lg py-4 px-6 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-lg">
              <GraduationCap size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">AutoComment VnEdu</h1>
              <p className="text-xs text-blue-100 font-medium">Bản nâng cấp Thông tư 27</p>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-2 bg-blue-700/50 px-3 py-1.5 rounded-full border border-blue-400/30 text-sm">
            <CheckCircle size={14} className="text-green-400" />
            <span>Chuẩn dữ liệu Học bạ</span>
          </div>
        </div>
      </header>

      <main className="flex-grow p-4 md:p-8 max-w-7xl mx-auto w-full flex flex-col gap-6">
        {/* Top Section */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Step 1: Upload */}
          <section className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6 flex flex-col transition-all hover:shadow-md">
            <div className="flex items-center gap-3 mb-4">
              <div className="bg-blue-50 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center font-bold">1</div>
              <h2 className="font-semibold text-gray-700">Tải bảng điểm Excel</h2>
            </div>
            
            {!originalFileName ? (
              <div 
                onClick={() => fileInputRef.current?.click()}
                className="flex-grow cursor-pointer border-2 border-dashed border-gray-200 rounded-xl bg-gray-50 flex flex-col items-center justify-center p-8 transition-all hover:border-blue-400 hover:bg-blue-50"
              >
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  className="hidden" 
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) => e.target.files?.[0] && handleFiles(e.target.files[0])}
                />
                <div className="bg-blue-100 p-4 rounded-full text-blue-600 mb-4">
                  <Upload size={32} />
                </div>
                <p className="text-sm font-medium text-gray-600">Chọn file từ máy tính</p>
                <p className="text-xs text-gray-400 mt-1">Hỗ trợ: .xlsx, .xls</p>
              </div>
            ) : (
              <div className="bg-green-50 border border-green-100 p-4 rounded-xl flex items-center justify-between">
                <div className="flex items-center gap-3 min-w-0">
                  <div className="bg-green-100 text-green-600 p-2 rounded-lg flex-shrink-0">
                    <FileSpreadsheet size={20} />
                  </div>
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-green-800 truncate">{originalFileName}</p>
                    <p className="text-xs text-green-600">{studentsData.length} học sinh</p>
                  </div>
                </div>
                <button onClick={reset} className="text-gray-400 hover:text-red-500 p-1">
                  <X size={18} />
                </button>
              </div>
            )}
          </section>

          {/* Step 2: Config */}
          <section className={cn(
            "bg-white rounded-2xl shadow-sm border border-gray-100 p-6 lg:col-span-2 flex flex-col gap-6 transition-all",
            !originalFileName && "opacity-50 pointer-events-none grayscale"
          )}>
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <div className="bg-blue-50 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center font-bold">2</div>
                <h2 className="font-semibold text-gray-700">Cấu hình & Tùy chọn</h2>
              </div>
              <button 
                onClick={processComments}
                className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2.5 rounded-xl font-semibold flex items-center justify-center gap-2 transition shadow-lg shadow-blue-200 active:scale-95 pulse-animation"
              >
                <Wand2 size={18} />
                Tạo nhận xét tự động
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Subject & Grade */}
              <div className="bg-gray-50/50 p-4 rounded-xl border border-gray-100 border-dashed flex flex-col gap-3">
                <div>
                  <label className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase mb-2">
                    <Book size={12} className="text-blue-500" />
                    Môn học
                  </label>
                  <select 
                    value={subject} 
                    onChange={(e) => setSubject(e.target.value)}
                    className="w-full bg-white border border-gray-200 rounded-lg px-3 py-2 text-sm font-medium focus:ring-2 focus:ring-blue-500 outline-none"
                  >
                    {Object.keys(ALL_SUBJECTS_BANK).map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div>
                  <label className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase mb-2">
                    <Users size={12} className="text-purple-500" />
                    Khối lớp
                  </label>
                  <div className="flex flex-wrap gap-2">
                    {["1", "2", "3", "4", "5", "general"].map(level => (
                      <button
                        key={level}
                        onClick={() => setGradeLevel(level)}
                        className={cn(
                          "px-2.5 py-1 rounded-md text-xs font-semibold border transition-all",
                          gradeLevel === level 
                            ? "bg-purple-600 border-purple-600 text-white shadow-sm"
                            : "bg-white border-gray-200 text-gray-500 hover:border-purple-200"
                        )}
                      >
                        {level === "general" ? "Chung" : `Lớp ${level}`}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Period & Toggles */}
              <div className="bg-gray-50/50 p-4 rounded-xl border border-gray-100 border-dashed flex flex-col gap-4">
                <div>
                  <label className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase mb-2">
                    <Calendar size={12} className="text-orange-500" />
                    Kỳ đánh giá (Đa chọn)
                  </label>
                  <div className="grid grid-cols-2 gap-2">
                    {(["GK1", "CK1", "GK2", "CK2"] as const).map((period) => (
                      <label 
                        key={period}
                        className={cn(
                          "flex items-center justify-between px-3 py-2 rounded-lg border cursor-pointer transition-all",
                          selectedPeriods.has(period) 
                            ? "bg-orange-50 border-orange-200 text-orange-700 shadow-sm" 
                            : "bg-white border-gray-100 text-gray-400 hover:border-orange-100"
                        )}
                      >
                        <span className="text-xs font-bold">{period}</span>
                        <input 
                          type="checkbox" 
                          className="w-3 h-3 rounded text-orange-500 border-gray-300 focus:ring-orange-400"
                          checked={selectedPeriods.has(period)}
                          onChange={(e) => {
                            const newSet = new Set(selectedPeriods);
                            if (e.target.checked) newSet.add(period);
                            else newSet.delete(period);
                            setSelectedPeriods(newSet);
                          }}
                        />
                      </label>
                    ))}
                  </div>
                </div>
                <div className="flex flex-col gap-2">
                  <label className="flex items-center gap-2 cursor-pointer group">
                    <input 
                      type="checkbox" 
                      checked={autoLevel} 
                      onChange={e => setAutoLevel(e.target.checked)}
                      className="w-4 h-4 rounded text-blue-600 border-gray-300 focus:ring-blue-500"
                    />
                    <span className="text-xs text-gray-600 font-medium group-hover:text-blue-600 transition-colors">Tự điền cột Mức dựa vào Điểm (Trừ GK)</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer group">
                    <input 
                      type="checkbox" 
                      className="w-4 h-4 rounded text-purple-600 border-gray-300 focus:ring-purple-500"
                      checked={addQualities}
                      onChange={e => setAddQualities(e.target.checked)}
                    />
                    <div className="flex flex-col">
                      <span className="text-xs text-gray-600 font-medium group-hover:text-purple-600 transition-colors">Ghép Năng lực & Phẩm chất</span>
                    </div>
                  </label>
                </div>
              </div>

              {/* Column Mapping */}
              <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex flex-col gap-2 relative">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-[10px] font-black text-blue-700 uppercase tracking-widest italic">Ánh xạ cột Excel</span>
                  <div className="flex items-center gap-1">
                    <button 
                      onClick={saveAsDefault}
                      className="flex items-center gap-1.5 px-2 py-1 bg-green-600 hover:bg-green-700 text-white text-[9px] font-bold rounded-md transition-all shadow-sm active:scale-95"
                      title="Lưu ánh xạ hiện tại làm mặc định"
                    >
                      <Save size={10} />
                      Lưu
                    </button>
                    <button 
                      onClick={loadDefault}
                      className="flex items-center gap-1.5 px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white text-[9px] font-bold rounded-md transition-all shadow-sm active:scale-95"
                      title="Tải cấu hình mặc định đã lưu"
                    >
                      <RotateCcw size={10} />
                      Tải mẫu
                    </button>
                    <button 
                      onClick={autoMap}
                      className="flex items-center gap-1.5 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-[9px] font-bold rounded-md transition-all shadow-sm active:scale-95"
                      title="Tự động nhận diện cột"
                    >
                      <RefreshCw size={10} />
                      Tự động
                    </button>
                    <button 
                      onClick={clearMappings}
                      className="flex items-center gap-1.5 px-2 py-1 bg-white hover:bg-red-50 text-red-500 border border-red-100 text-[9px] font-bold rounded-md transition-all shadow-sm active:scale-95"
                      title="Xóa tất cả ánh xạ"
                    >
                      <Trash2 size={10} />
                      Xóa
                    </button>
                  </div>
                </div>
                <div className="space-y-1.5 overflow-y-auto pr-1">
                  <div className="grid grid-cols-2 gap-2 mb-2 pb-2 border-b border-blue-100">
                    <div className="flex flex-col gap-1">
                      <span className="text-[9px] font-bold text-blue-400 uppercase">Họ Tên</span>
                      <select 
                        value={colMapping.name} 
                        onChange={e => setColMapping(prev => ({...prev, name: e.target.value}))}
                        className="w-full bg-white border border-blue-200 rounded p-1 text-[11px] font-medium outline-none"
                      >
                        <option value="">-- Chọn --</option>
                        {headers.map(h => <option key={h} value={h}>{h}</option>)}
                      </select>
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-[9px] font-bold text-blue-400 uppercase">Mã HS</span>
                      <select 
                        value={colMapping.studentId} 
                        onChange={e => setColMapping(prev => ({...prev, studentId: e.target.value}))}
                        className="w-full bg-white border border-blue-200 rounded p-1 text-[11px] font-medium outline-none"
                      >
                        <option value="">-- Không --</option>
                        {headers.map(h => <option key={h} value={h}>{h}</option>)}
                      </select>
                    </div>
                  </div>

                  {/* Period Mappings */}
                  <div className="flex flex-col gap-3 max-h-[300px] overflow-y-auto pb-4">
                    {(["GK1", "CK1", "GK2", "CK2"] as const).filter(p => selectedPeriods.has(p)).map(period => (
                      <div key={period} className="bg-white/60 p-2 rounded-lg border border-blue-200/50">
                        <div className="flex items-center gap-2 mb-2">
                          <div className="w-1.5 h-1.5 rounded-full bg-orange-400" />
                          <span className="text-[10px] font-black text-orange-600">{period}</span>
                        </div>
                        <div className="grid grid-cols-3 gap-2">
                          <div className="flex flex-col gap-1">
                            <span className="text-[8px] font-bold text-gray-400 uppercase">Điểm</span>
                            <select 
                              value={colMapping.periods[period].score} 
                              onChange={e => {
                                const newMap = { ...colMapping };
                                newMap.periods[period].score = e.target.value;
                                setColMapping(newMap);
                              }}
                              className="w-full bg-white border border-gray-100 rounded p-1 text-[10px] outline-none"
                            >
                              <option value="">--</option>
                              {headers.map(h => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </div>
                          <div className="flex flex-col gap-1">
                            <span className="text-[8px] font-bold text-gray-400 uppercase">Mức</span>
                            <select 
                              value={colMapping.periods[period].level} 
                              onChange={e => {
                                const newMap = { ...colMapping };
                                newMap.periods[period].level = e.target.value;
                                setColMapping(newMap);
                              }}
                              className="w-full bg-white border border-gray-100 rounded p-1 text-[10px] outline-none"
                            >
                              <option value="">--</option>
                              {headers.map(h => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </div>
                          <div className="flex flex-col gap-1">
                            <span className="text-[8px] font-bold text-blue-700 uppercase">N.Xét</span>
                            <select 
                              value={colMapping.periods[period].comment} 
                              onChange={e => {
                                const newMap = { ...colMapping };
                                newMap.periods[period].comment = e.target.value;
                                setColMapping(newMap);
                              }}
                              className="w-full bg-blue-50 border border-blue-300 rounded p-1 text-[10px] font-bold text-blue-800 outline-none"
                            >
                              <option value="">--</option>
                              {headers.map(h => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </section>
        </div>

        {/* Data Table Section */}
        <AnimatePresence>
          {studentsData.length > 0 && (
            <motion.section 
              initial={{opacity: 0, y: 20}}
              animate={{opacity: 1, y: 0}}
              exit={{opacity: 0, y: 20}}
              className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden flex flex-col flex-grow min-h-[500px]"
            >
              <div className="p-4 bg-gray-50 border-b border-gray-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <div className="bg-blue-100 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center font-bold">3</div>
                  <h2 className="font-semibold text-gray-700">Dữ liệu học sinh</h2>
                  <span className="bg-blue-100 text-blue-700 text-[10px] font-bold px-2 py-0.5 rounded-full border border-blue-200">
                    {studentsData.length} HS
                  </span>
                  <label className="flex items-center gap-2 cursor-pointer ml-4 group">
                    <input 
                      type="checkbox" 
                      checked={filterEmptyComments} 
                      onChange={e => setFilterEmptyComments(e.target.checked)}
                      className="w-4 h-4 rounded text-blue-600 border-gray-300 focus:ring-blue-500"
                    />
                    <span className="text-xs text-gray-500 group-hover:text-gray-700 transition-colors">Ẩn HS chưa có NX</span>
                  </label>
                </div>
                <div className="flex items-center gap-2 w-full sm:w-auto">
                  <button 
                    onClick={copyAll}
                    disabled={!isProcessed}
                    className="flex-1 sm:flex-initial flex items-center justify-center gap-2 bg-white border border-gray-200 px-4 py-2 rounded-xl text-sm font-semibold hover:bg-gray-100 disabled:opacity-50 transition-all"
                  >
                    <Copy size={16} className="text-blue-500" />
                    Copy
                  </button>
                  <button 
                    onClick={exportCommentsOnly}
                    disabled={!isProcessed}
                    className="flex-1 sm:flex-initial flex items-center justify-center gap-2 bg-white border border-blue-200 text-blue-700 px-4 py-2 rounded-xl text-sm font-semibold hover:bg-blue-50 disabled:opacity-50 transition-all"
                  >
                    <FileSpreadsheet size={16} />
                    Xuất riêng nhận xét
                  </button>
                  <button 
                    onClick={exportExcel}
                    className="flex-1 sm:flex-initial flex items-center justify-center gap-2 bg-green-600 text-white px-4 py-2 rounded-xl text-sm font-semibold hover:bg-green-700 transition shadow-lg shadow-green-100 active:scale-95"
                  >
                    <Download size={16} />
                    Lưu file
                  </button>
                </div>
              </div>

              <div className="overflow-auto table-container flex-grow relative">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-gray-50/80 backdrop-blur sticky top-0 z-10">
                    <tr>
                      <th className="px-6 py-4 border-b border-gray-100 font-bold text-gray-400 text-[10px] uppercase tracking-wider w-12 text-center">STT</th>
                      <th className="px-6 py-4 border-b border-gray-100 font-bold text-gray-400 text-[10px] uppercase tracking-wider w-40">Họ và Tên</th>
                      {((["GK1", "CK1", "GK2", "CK2"] as const).filter(p => selectedPeriods.has(p))).map(period => (
                        <th key={period} className="px-4 py-4 border-b border-gray-100 font-bold text-gray-400 text-[10px] uppercase tracking-wider min-w-[200px]">
                          Nhận xét {period}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {studentsData.map((s, idx) => {
                      const name = s[colMapping.name] || "";
                      if (String(name).toLowerCase().includes('tổng')) return null;
                      
                      // Skip if filtering empty and no comments in ANY selected period
                      if (filterEmptyComments) {
                        const hasAnyComment = Array.from(selectedPeriods).some(period => {
                          const pMap = colMapping.periods[period];
                          return (s[pMap.comment] || "").trim().length > 0;
                        });
                        if (!hasAnyComment) return null;
                      }

                      return (
                        <tr key={idx} className="hover:bg-blue-50/30 transition-colors group">
                          <td className="px-6 py-4 text-gray-400 font-medium text-center">{idx + 1}</td>
                          <td className="px-6 py-4 font-semibold text-gray-700">{name}</td>
                          {((["GK1", "CK1", "GK2", "CK2"] as const).filter(p => selectedPeriods.has(p))).map(period => {
                            const pMap = colMapping.periods[period];
                            const level = s[pMap.level];
                            const levelColor = level === 'T' ? 'bg-green-100 text-green-700 border-green-200' : (level === 'H' ? 'bg-yellow-100 text-yellow-700 border-yellow-200' : (level === 'C' ? 'bg-red-100 text-red-700 border-red-200' : 'bg-gray-100 text-gray-500 border-gray-200'));

                            return (
                              <td key={period} className="px-4 py-4">
                                <div className="flex flex-col gap-2">
                                  <div className="flex items-center gap-2">
                                    <span className="text-[10px] font-black text-gray-400">{period}:</span>
                                    {s[pMap.score] && <span className="text-[10px] font-bold text-blue-600">Đ: {s[pMap.score]}</span>}
                                    {level && <span className={cn("px-1.5 py-0.5 rounded text-[9px] font-bold border", levelColor)}>{level}</span>}
                                  </div>
                                  <textarea
                                    value={s[pMap.comment] || ""}
                                    onChange={(e) => {
                                      const newS = [...studentsData];
                                      newS[idx] = {...newS[idx], [pMap.comment]: e.target.value};
                                      setStudentsData(newS);
                                    }}
                                    placeholder={`Nhận xét ${period}...`}
                                    className={cn(
                                      "w-full bg-white border border-gray-100 rounded-lg p-2 text-[11px] outline-none transition-all focus:ring-1 focus:ring-blue-400 focus:border-blue-400 h-16 resize-none",
                                      !s[pMap.comment] && "text-gray-400 italic"
                                    )}
                                  />
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </motion.section>
          )}
        </AnimatePresence>
      </main>

      {/* Confirmation Dialog */}
      <AnimatePresence>
        {showConfirm && (
          <>
            <motion.div 
              initial={{opacity: 0}}
              animate={{opacity: 1}}
              exit={{opacity: 0}}
              onClick={() => setShowConfirm(null)}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[60]"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, x: "-50%", y: "-40%" }}
              animate={{ opacity: 1, scale: 1, x: "-50%", y: "-50%" }}
              exit={{ opacity: 0, scale: 0.8, x: "-50%", y: "-40%" }}
              transition={{ type: "spring", damping: 25, stiffness: 300 }}
              className="fixed top-1/2 left-1/2 z-[70] bg-white rounded-3xl shadow-2xl p-8 w-[90%] max-w-md border border-gray-100"
            >
              <motion.div 
                initial={{ scale: 0.5, rotate: -10 }}
                animate={{ scale: 1, rotate: 0 }}
                transition={{ delay: 0.1, type: "spring" }}
                className="bg-blue-100 text-blue-600 w-16 h-16 rounded-2xl flex items-center justify-center mb-6 mx-auto"
              >
                <Download size={32} />
              </motion.div>
              <motion.h3 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.2 }}
                className="text-xl font-bold text-gray-800 text-center mb-2"
              >
                Bạn muốn xuất file?
              </motion.h3>
              <motion.p 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.25 }}
                className="text-sm text-gray-500 text-center mb-8 leading-relaxed"
              >
                {showConfirm.type === 'all' 
                  ? "Hệ thống sẽ ghi đè nhận xét vào file Excel gốc của bạn. Hãy đảm bảo bạn đã kiểm tra kỹ nội dung."
                  : "Hệ thống sẽ tạo một file Excel mới chỉ chứa danh sách nhận xét đã tạo."}
              </motion.p>
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.3 }}
                className="flex gap-4"
              >
                <button 
                  onClick={() => setShowConfirm(null)}
                  className="flex-1 px-6 py-3 rounded-xl font-bold bg-gray-100 text-gray-600 hover:bg-gray-200 transition-all active:scale-95"
                >
                  Hủy
                </button>
                <button 
                  onClick={showConfirm.action}
                  className="flex-1 px-6 py-3 rounded-xl font-bold bg-blue-600 text-white hover:bg-blue-700 shadow-lg shadow-blue-200 transition-all active:scale-95"
                >
                  Đồng ý
                </button>
              </motion.div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Toast Notification */}
      <AnimatePresence>
        {toast && (
          <motion.div 
            initial={{opacity: 0, y: 50, scale: 0.9}}
            animate={{opacity: 1, y: 0, scale: 1}}
            exit={{opacity: 0, y: 50, scale: 0.9}}
            className={cn(
              "fixed bottom-8 left-1/2 -translate-x-1/2 z-50 flex items-center gap-3 px-6 py-3 rounded-2xl shadow-2xl border backdrop-blur-sm",
              toast.type === 'success' ? "bg-gray-900/90 text-white border-gray-800" : "bg-red-600 text-white border-red-500"
            )}
          >
            {toast.type === 'success' ? <FileCheck size={20} className="text-green-400" /> : <X size={20} />}
            <span className="text-sm font-semibold">{toast.msg}</span>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
