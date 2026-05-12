import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileDown, Settings, Users, CheckCircle2, AlertCircle, Info, Pencil, X, Save } from 'lucide-react';

const defaultCommentBank = `Năng lực/Phẩm chất,Mức,Nhận xét
Yêu nước,Tốt,"Em yêu thiên nhiên, quê hương và có những việc làm thiết thực bảo vệ thiên nhiên."
Yêu nước,Tốt,"Em tích cực tham gia các hoạt động tập thể và thực hiện tốt nội quy trường lớp."
Yêu nước,Tốt,"Em biết trân trọng, bảo vệ tự nhiên và tự hào về truyền thống quê hương."
Yêu nước,Đạt,"Em có ý thức thực hiện nội quy và tham gia các hoạt động chung."
Yêu nước,Đạt,"Em biết yêu quê hương đất nước qua các bài học."
Yêu nước,Đạt,"Em tham gia đầy đủ các hoạt động của trường lớp."
Yêu nước,Cần cố gắng,"Em cần tích cực hơn trong các hoạt động tập thể."
Yêu nước,Cần cố gắng,"Em cần có ý thức hơn trong việc bảo vệ môi trường và của công."
Yêu nước,Cần cố gắng,"Em ít tham gia các phong trào của lớp, cần cố gắng hơn."
Nhân ái,Tốt,"Em hòa đồng, biết giúp đỡ và chia sẻ với bạn bè."
Nhân ái,Tốt,"Em vui vẻ hòa đồng, tôn trọng người lớn tuổi, yêu quý bạn bè."
Nhân ái,Tốt,"Em có tấm lòng nhân ái, sẵn sàng giúp đỡ người gặp hoàn cảnh khó khăn."
Nhân ái,Đạt,"Em cư xử hòa nhã với bạn bè và thầy cô."
Nhân ái,Đạt,"Em biết yêu thương và giúp đỡ bạn bè phù hợp với hoàn cảnh."
Nhân ái,Đạt,"Em biết chia sẻ buồn vui cùng bạn bè trong lớp."
Nhân ái,Cần cố gắng,"Em cần chủ động hơn trong việc giúp đỡ bạn bè."
Nhân ái,Cần cố gắng,"Em còn hay trêu chọc bạn, cần đoàn kết hơn."
Nhân ái,Cần cố gắng,"Em chưa biết kiềm chế cảm xúc, cần hòa nhã hơn với bạn."
Chăm chỉ,Tốt,"Em đi học đầy đủ, đúng giờ, tự giác thực hiện nhiệm vụ học tập."
Chăm chỉ,Tốt,"Em chăm học và hoàn thành xuất sắc nhiệm vụ học tập."
Chăm chỉ,Tốt,"Em có ý thức học hỏi, tự giác làm bài và có ý chí vượt khó."
Chăm chỉ,Đạt,"Em có ý thức học tập và hoàn thành nhiệm vụ được giao."
Chăm chỉ,Đạt,"Em chăm chú nghe giảng và làm bài tập đầy đủ."
Chăm chỉ,Đạt,"Em hoàn thành các bài học trên lớp."
Chăm chỉ,Cần cố gắng,"Em cần tập trung hơn trong học tập."
Chăm chỉ,Cần cố gắng,"Em chưa chăm học đều, thường xuyên quên sách vở."
Chăm chỉ,Cần cố gắng,"Em hay làm việc riêng trong giờ học, cần chú ý hơn."
Trung thực,Tốt,"Em trung thực trong học tập và sinh hoạt, không tham của rơi."
Trung thực,Tốt,"Em luôn thật thà, biết nhận lỗi và sửa lỗi khi phạm lỗi."
Trung thực,Tốt,"Em không nói dối, trung thực trong kiểm tra và đánh giá."
Trung thực,Đạt,"Em thực hiện tương đối tốt các yêu cầu học tập, biết nhận lỗi."
Trung thực,Đạt,"Em có ý thức nói thật, không lấy đồ của bạn."
Trung thực,Đạt,"Em biết đồng tình with những hành vi trung thực."
Trung thực,Cần cố gắng,"Em cần mạnh dạn và trung thực hơn trong học tập."
Trung thực,Cần cố gắng,"Em chưa tự giác nhận lỗi khi làm sai."
Trung thực,Cần cố gắng,"Em còn chưa thật thà trong một số hoạt động, cần khắc phục."
Trách nhiệm,Tốt,"Em có trách nhiệm với công việc được giao ở trường, lớp."
Trách nhiệm,Tốt,"Em có ý thức bảo vệ của công, giữ gìn vệ sinh môi trường sạch sẽ."
Trách nhiệm,Tốt,"Em luôn hoàn thành xuất sắc các nhiệm vụ do nhóm phân công."
Trách nhiệm,Đạt,"Em thực hiện tương đối đầy đủ nhiệm vụ được giao."
Trách nhiệm,Đạt,"Em có tham gia trực nhật và giữ vệ sinh lớp."
Trách nhiệm,Đạt,"Em biết bảo quản đồ dùng học tập của bản thân."
Trách nhiệm,Cần cố gắng,"Em cần nâng cao tinh thần trách nhiệm với tập thể."
Trách nhiệm,Cần cố gắng,"Em hay quên trực nhật, cần có ý thức hơn."
Trách nhiệm,Cần cố gắng,"Em chưa biết bảo vệ của chung, cần nhắc nhở nhiều."
Tự chủ và tự học,Tốt,"Em có ý thức tự học và tự chủ trong mọi vấn đề."
Tự chủ và tự học,Tốt,"Em có khả năng tự thực hiện tốt các nhiệm vụ học tập."
Tự chủ và tự học,Tốt,"Em tự giác làm bài không cần nhắc nhở, biết tự điều chỉnh hành vi."
Tự chủ và tự học,Đạt,"Em bước đầu biết tự thực hiện nhiệm vụ học tập."
Tự chủ và tự học,Đạt,"Em biết tự phục vụ nhu cầu bản thân ở mức cơ bản."
Tự chủ và tự học,Đạt,"Em hoàn thành bài tập khi có sự hướng dẫn của giáo viên."
Tự chủ và tự học,Cần cố gắng,"Em chưa có ý thức tự học, cần nhắc nhở nhiều."
Tự chủ và tự học,Cần cố gắng,"Em phụ thuộc vào người khác, chưa tự giác làm bài."
Tự chủ và tự học,Cần cố gắng,"Em cần chủ động hơn trong học tập và rèn luyện."
Giao tiếp và hợp tác,Tốt,"Em trình bày rõ ràng, ngắn gọn nội dung cần trao đổi."
Giao tiếp và hợp tác,Tốt,"Em biết trao đổi ý kiến cùng bạn rất tốt, mạnh dạn đưa ra ý kiến cá nhân."
Giao tiếp và hợp tác,Tốt,"Em có khả năng phối hợp với bạn bè khi làm việc nhóm rất hiệu quả."
Giao tiếp và hợp tác,Đạt,"Em biết cách giao tiếp cơ bản và hợp tác với nhóm."
Giao tiếp và hợp tác,Đạt,"Em biết lắng nghe và chia sẻ ý kiến với các bạn."
Giao tiếp và hợp tác,Đạt,"Em tham gia hoạt động nhóm tương đối tốt."
Giao tiếp và hợp tác,Cần cố gắng,"Em chưa mạnh dạn trong giao tiếp và hợp tác."
Giao tiếp và hợp tác,Cần cố gắng,"Em còn thụ động, ít phát biểu, cần cố gắng hơn."
Giao tiếp và hợp tác,Cần cố gắng,"Em ít tham gia vào các hoạt động thảo luận nhóm."
Giải quyết vấn đề,Tốt,"Em biết thu thập thông tin và tình huống, nhận ra những vấn đề đơn giản."
Giải quyết vấn đề,Tốt,"Em có năng lực giải quyết vấn đề, sáng tạo trong học tập."
Giải quyết vấn đề,Tốt,"Em xử lý tình huống nhanh, chính xác và có ý tưởng độc đáo."
Giải quyết vấn đề,Đạt,"Em có khả năng giải quyết các vấn đề cơ bản."
Giải quyết vấn đề,Đạt,"Em bước đầu biết cách thức giải quyết vấn đề theo hướng dẫn."
Giải quyết vấn đề,Đạt,"Em biết xử lý các tình huống quen thuộc trong bài học."
Giải quyết vấn đề,Cần cố gắng,"Em xử lí tình huống còn lúng túng, chưa linh hoạt."
Giải quyết vấn đề,Cần cố gắng,"Em chưa biết cách giải quyết các bài tập khó, cần nỗ lực hơn."
Giải quyết vấn đề,Cần cố gắng,"Em cần mạnh dạn hơn khi đối mặt với tình huống mới."
Ngôn ngữ,Tốt,"Đọc to, rõ ràng, lưu loát. Câu văn ngắn gọn, dễ hiểu."
Ngôn ngữ,Tốt,"Viết chính tả chính xác. Chữ viết đều, đúng nét, trình bày sạch đẹp."
Ngôn ngữ,Tốt,"Em có vốn từ phong phú, diễn đạt câu trọn vẹn, sinh động."
Ngôn ngữ,Đạt,"Đọc bài khá lưu loát, chữ viết tương đối rõ ràng."
Ngôn ngữ,Đạt,"Hiểu nội dung bài đọc và hoàn thành tốt các bài tập."
Ngôn ngữ,Đạt,"Em có kỹ năng nghe, nói, đọc, viết ở mức đạt yêu cầu."
Ngôn ngữ,Cần cố gắng,"Đọc còn đánh vần chậm, mắc nhiều lỗi chính tả."
Ngôn ngữ,Cần cố gắng,"Kĩ năng đọc hiểu văn bản còn hạn chế."
Ngôn ngữ,Cần cố gắng,"Chữ viết chưa đúng mẫu, trình bày bài chưa khoa học."
Tính toán,Tốt,"Nắm vững kiến thức đã học. Tính toán thành thạo, giải toán đúng."
Tính toán,Tốt,"Thông minh, có trí nhớ tốt, tính toán nhanh và chính xác."
Tính toán,Tốt,"Có tư duy toán tốt, giải toán có lời văn rất sáng tạo."
Tính toán,Đạt,"Hiểu bài, nắm được các kiến thức cơ bản."
Tính toán,Đạt,"Thực hiện được các phép tính nhưng giải toán có lời văn còn chậm."
Tính toán,Đạt,"Đã hoàn thành nội dung môn học nhưng tính toán đôi khi nhầm lẫn."
Tính toán,Cần cố gắng,"Chưa nắm chắc kiến thức, kĩ năng tính toán còn hạn chế."
Tính toán,Cần cố gắng,"Tiếp thu bài chậm, tính toán chưa cẩn thận."
Tính toán,Cần cố gắng,"Chưa nắm được cách giải toán có lời văn, cần rèn luyện nhiều."
Khoa học,Tốt,"Hiểu bài, vận dụng tốt các kiến thức đã học vào cuộc sống."
Khoa học,Tốt,"Chủ động nắm bắt và ghi nhớ kiến thức tốt, ham tìm tòi khám phá."
Khoa học,Tốt,"Có vốn hiểu biết phong phú về tự nhiên và xã hội."
Khoa học,Đạt,"Hoàn thành nội dung kiến thức môn học."
Khoa học,Đạt,"Nắm được kiến thức cơ bản và biết vận dụng bài học vào thực tiễn."
Khoa học,Đạt,"Biết kể tên và nhận diện các sự vật hiện tượng xung quanh."
Khoa học,Cần cố gắng,"Cần tích cực chủ động trong học tập hơn."
Khoa học,Cần cố gắng,"Chưa hoàn thành kiến thức kĩ năng của môn học."
Khoa học,Cần cố gắng,"Vốn hiểu biết về tự nhiên còn hạn chế."
Thẩm mĩ,Tốt,"Có khiếu thẩm mĩ, thực hành gấp cắt dán sản phẩm đẹp, sáng tạo."
Thẩm mĩ,Tốt,"Hát hay, biểu diễn tự nhiên, tự tin thể hiện sắc thái bài hát."
Thẩm mĩ,Tốt,"Em luôn hoàn thành bài vẽ nhanh và có nhiều ý tưởng độc đáo."
Thẩm mĩ,Đạt,"Hoàn thành sản phẩm đạt yêu cầu."
Thẩm mĩ,Đạt,"Thuộc lời ca, hát đúng giai điệu bài hát."
Thẩm mĩ,Đạt,"Biết phối hợp màu sắc cơ bản trong các bài vẽ."
Thẩm mĩ,Cần cố gắng,"Cần nhanh nhẹn hơn trong các thao tác kĩ thuật."
Thẩm mĩ,Cần cố gắng,"Hát còn nhỏ, chưa đúng giai điệu, cần luyện tập thêm."
Thẩm mĩ,Cần cố gắng,"Sản phẩm mĩ thuật chưa hoàn thiện, cần cố gắng."
Thể chất,Tốt,"Hoàn thành tốt lượng vận động, tham gia tích cực các trò chơi."
Thể chất,Tốt,"Thực hiện chuẩn xác các động tác thể dục, nhanh nhẹn, hoạt bát."
Thể chất,Tốt,"Em tự giác tập luyện thể dục thể thao, phát huy tốt năng khiếu."
Thể chất,Đạt,"Em đã thực hiện cơ bản nội dung, động tác đã học."
Thể chất,Đạt,"Có ý thức rèn luyện, tham gia các trò chơi vận động."
Thể chất,Đạt,"Hoàn thành lượng vận động của bài tập."
Thể chất,Cần cố gắng,"Chưa tích cực tham gia vận động, hay đùa nghịch trong giờ."
Thể chất,Cần cố gắng,"Chưa hoàn thành kỹ thuật động tác, cần cố gắng hơn."
Thể chất,Cần cố gắng,"Thể lực còn yếu, cần thường xuyên tập luyện thể thao."
Công nghệ,Tốt,"Em có hiểu biết tốt về công nghệ và biết vận dụng thực hành."
Công nghệ,Tốt,"Hoàn thành sản phẩm kĩ thuật nhanh, đẹp và có tính ứng dụng cao."
Công nghệ,Tốt,"Có khả năng lắp ráp, thao tác các mô hình công nghệ xuất sắc."
Công nghệ,Đạt,"Em có kiến thức cơ bản về môn công nghệ."
Công nghệ,Đạt,"Sản phẩm thực hành hoàn thành đạt yêu cầu của bài học."
Công nghệ,Đạt,"Biết sử dụng các dụng cụ an toàn trong giờ học."
Công nghệ,Cần cố gắng,"Em cần chú ý tìm hiểu thêm kiến thức về công nghệ."
Công nghệ,Cần cố gắng,"Thao tác lắp ráp còn lóng ngóng, chưa hoàn thiện sản phẩm."
Công nghệ,Cần cố gắng,"Cần cẩn thận hơn khi sử dụng các dụng cụ thực hành."
Tin học,Tốt,"Em thao tác tốt với chuột và bàn phím, sử dụng máy tính an toàn."
Tin học,Tốt,"Thực hiện nhanh và chính xác các thao tác cơ bản trên phần mềm."
Tin học,Tốt,"Có năng khiếu tin học, tiếp thu nhanh các lệnh trên máy tính."
Tin học,Đạt,"Em biết thực hiện được các thao tác cơ bản với máy tính."
Tin học,Đạt,"Biết khởi động, thoát phần mềm đúng cách."
Tin học,Đạt,"Hoàn thành các bài tập thực hành trên máy ở mức cơ bản."
Tin học,Cần cố gắng,"Em cần thực hành nhiều hơn các thao tác máy tính."
Tin học,Cần cố gắng,"Chưa quen với thao tác chuột và bàn phím."
Tin học,Cần cố gắng,"Chưa nắm được các bước khởi động và sử dụng phần mềm."`;

const criteriaMapping: Record<string, string[]> = {
  'Yêu nước': ['yêu nước'],
  'Nhân ái': ['nhân ái'],
  'Chăm chỉ': ['chăm chỉ'],
  'Trung thực': ['trung thực'],
  'Trách nhiệm': ['trách nhiệm'],
  'Tự chủ và tự học': ['tự chủ', 'tự học'],
  'Giao tiếp và hợp tác': ['giao tiếp', 'hợp tác'],
  'Giải quyết vấn đề': ['giải quyết vấn đề', 'gqvđ', 'sáng tạo'],
  'Ngôn ngữ': ['ngôn ngữ'],
  'Tính toán': ['tính toán'],
  'Khoa học': ['khoa học'],
  'Thẩm mĩ': ['thẩm mĩ', 'thẩm mỹ'],
  'Thể chất': ['thể chất'],
  'Công nghệ': ['công nghệ'],
  'Tin học': ['tin học']
};

const categoryMap: Record<string, string> = {
  'Tự chủ và tự học': 'chung',
  'Giao tiếp và hợp tác': 'chung',
  'Giải quyết vấn đề': 'chung',
  'Ngôn ngữ': 'dacThu',
  'Tính toán': 'dacThu',
  'Khoa học': 'dacThu',
  'Thẩm mĩ': 'dacThu',
  'Thể chất': 'dacThu',
  'Công nghệ': 'dacThu',
  'Tin học': 'dacThu',
  'Yêu nước': 'phamChat',
  'Nhân ái': 'phamChat',
  'Chăm chỉ': 'phamChat',
  'Trung thực': 'phamChat',
  'Trách nhiệm': 'phamChat'
};

const levelMapping: Record<string, string> = {
  'T': 'Tốt',
  'Đ': 'Đạt',
  'C': 'Cần cố gắng',
  'TỐT': 'Tốt',
  'ĐẠT': 'Đạt',
  'CẦN CỐ GẮNG': 'Cần cố gắng',
  'HTT': 'Tốt',
  'H': 'Đạt',
  'CHT': 'Cần cố gắng'
};

function parseCSVRow(text: string) {
  let ret = [''], i = 0, s = true;
  for (let l = text.length; i < l; i++) {
    let c = text[i];
    if (c === '"') {
      s = !s;
    } else if (c === ',' && s) {
      ret.push('');
    } else {
      ret[ret.length - 1] += c;
    }
  }
  return ret.map(val => val.replace(/^"|"$/g, '').trim());
}

export default function CompetencyApp() {
  const [activeTab, setActiveTab] = useState('generator');
  const [commentBankText, setCommentBankText] = useState(defaultCommentBank);
  const [parsedBank, setParsedBank] = useState<Record<string, Record<string, string[]>>>({});
  const [students, setStudents] = useState<any[]>([]);
  const [fileName, setFileName] = useState('');
  const [statusMsg, setStatusMsg] = useState({ text: '', type: '' });
  const [gradeBlock, setGradeBlock] = useState('k12'); 
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [editingStudent, setEditingStudent] = useState<any>(null);

  useEffect(() => {
    const lines = commentBankText.split('\n');
    const bank: Record<string, Record<string, string[]>> = {};
    Object.keys(criteriaMapping).forEach(c => { bank[c] = {}; });

    lines.slice(1).forEach(line => {
      if (!line.trim()) return;
      const [criteria, level, ...commentParts] = parseCSVRow(line);
      let comment = commentParts.join(',').replace(/^"|"$/g, '').trim();
      
      if (comment && comment.length > 0) {
        if (comment === comment.toUpperCase()) {
          comment = comment.charAt(0).toUpperCase() + comment.slice(1).toLowerCase();
        } else {
          comment = comment.charAt(0).toUpperCase() + comment.slice(1);
        }
      }
      
      if (criteria && level && bank[criteria]) {
        if (!bank[criteria][level]) {
          bank[criteria][level] = [];
        }
        bank[criteria][level].push(comment);
      }
    });
    setParsedBank(bank);
  }, [commentBankText]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setStudents([]);
    setStatusMsg({ text: 'Đang tải và phân tích dữ liệu...', type: 'info' });
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false }) as any[][];
        processDataRows(rows);
      } catch (error) {
        setStatusMsg({ text: "Lỗi khi đọc nội dung file.", type: 'error' });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processDataRows = (rows: any[][]) => {
    if (rows.length === 0) {
      setStatusMsg({ text: "File rỗng.", type: 'error' });
      return;
    }

    let colIndices = { name: -1, studentId: -1 };
    let foundCriteriaCols: Record<string, number> = {};
    let dataStartIndex = -1;

    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const cols = rows[i].map(c => String(c).toLowerCase().trim());
      
      cols.forEach((colStr, index) => {
        if (colStr.includes('họ và tên') || colStr.includes('họ tên') || colStr === 'tên học sinh') colIndices.name = index;
        if (colStr.includes('mã học sinh') || colStr.includes('mã hs') || colStr.includes('mã định danh')) colIndices.studentId = index;

        Object.keys(criteriaMapping).forEach(mainCriteria => {
          const keywords = criteriaMapping[mainCriteria];
          if (keywords.some(kw => colStr.includes(kw))) {
            foundCriteriaCols[mainCriteria] = index;
          }
        });
      });

      if (cols.length > 0 && String(cols[0]).trim() !== '' && /^\d+$/.test(String(cols[0]).trim())) {
        if (dataStartIndex === -1) dataStartIndex = i;
      }
    }

    if (colIndices.name === -1 || Object.keys(foundCriteriaCols).length === 0) {
      setStatusMsg({ text: "Không tìm thấy cột Họ Tên hoặc các cột Năng lực/Phẩm chất. Vui lòng kiểm tra file.", type: 'error' });
      return;
    }

    const results = [];
    const startIdx = dataStartIndex !== -1 ? dataStartIndex : 10;

    for (let i = startIdx; i < rows.length; i++) {
      const row = rows[i];
      if (row.length <= colIndices.name || !row[colIndices.name]) continue;

      const name = String(row[colIndices.name]).trim();
      if (!name || name.toLowerCase().includes('tổng số') || name.toLowerCase().includes('trung bình')) continue;
      
      const studentId = colIndices.studentId !== -1 ? String(row[colIndices.studentId]).trim() : '';

      const originalGrades: Record<string, string> = {};
      const individualComments: Record<string, string> = {}; 
      const studentIndex = results.length; 
      
      Object.keys(foundCriteriaCols).forEach(criteria => {
        const colIdx = foundCriteriaCols[criteria];
        let rawGrade = row[colIdx] ? String(row[colIdx]).trim() : '';
        let gradeStr = rawGrade.toUpperCase();
        originalGrades[criteria] = gradeStr;
        const level = levelMapping[gradeStr] || gradeStr;
        
        const possibleComments = parsedBank[criteria] && parsedBank[criteria][level];
        if (possibleComments && possibleComments.length > 0) {
          individualComments[criteria] = possibleComments[studentIndex % possibleComments.length];
        } else {
          let fallback = rawGrade;
          if (fallback && fallback.length > 0) {
            if (fallback === fallback.toUpperCase()) {
              fallback = fallback.charAt(0).toUpperCase() + fallback.slice(1).toLowerCase();
            } else {
              fallback = fallback.charAt(0).toUpperCase() + fallback.slice(1);
            }
          }
          individualComments[criteria] = fallback;
        }
      });

      results.push({
        id: `std_${i}`,
        studentId,
        name,
        originalGrades,
        individualComments,
        comments: generateGroupedComments(individualComments)
      });
    }

    setStudents(results);
    setStatusMsg({ text: `Xử lý thành công! Đã quy chuẩn và đa dạng hoá lời nhận xét cho ${results.length} học sinh.`, type: 'success' });
  };

  const generateGroupedComments = (individualComments: Record<string, string>) => {
    let grouped: Record<string, string[]> = { chung: [], dacThu: [], phamChat: [] };
    
    Object.keys(individualComments).forEach(criteria => {
      const comment = individualComments[criteria];
      const cat = categoryMap[criteria];
      
      if (gradeBlock === 'k12' && (criteria === 'Công nghệ' || criteria === 'Tin học')) return;

      if (cat && comment && comment.length > 2) {
        grouped[cat].push(comment);
      }
    });

    const formatString = (arr: string[]) => {
      if (arr.length === 0) return "";
      let str = arr.join(' ');
      return str.charAt(0).toUpperCase() + str.slice(1);
    };

    return {
      chung: formatString(grouped.chung),
      dacThu: formatString(grouped.dacThu),
      phamChat: formatString(grouped.phamChat)
    };
  };

  const openEditModal = (student: any) => {
    setEditingStudent(JSON.parse(JSON.stringify(student)));
  };

  const closeEditModal = () => {
    setEditingStudent(null);
  };

  const handleEditIndividualComment = (criteria: string, value: string) => {
    setEditingStudent((prev: any) => {
      const updatedIndividual = { ...prev.individualComments, [criteria]: value };
      return {
        ...prev,
        individualComments: updatedIndividual
      };
    });
  };

  const handleEditGroupedComment = (cat: string, value: string) => {
    setEditingStudent((prev: any) => ({
      ...prev,
      comments: { ...prev.comments, [cat]: value }
    }));
  };

  const regenerateGroupedForEditingStudent = () => {
    if (!editingStudent) return;
    const newGrouped = generateGroupedComments(editingStudent.individualComments);
    setEditingStudent((prev: any) => ({
      ...prev,
      comments: newGrouped
    }));
  };

  const saveEditedStudent = () => {
    setStudents(prevStudents => 
      prevStudents.map(std => std.id === editingStudent.id ? editingStudent : std)
    );
    closeEditModal();
  };

  const exportToStandardTemplateExcel = async () => {
    if (students.length === 0) return;
    try {
      let wsData = [];
      let merges = [];
      let cols = [];

      if (gradeBlock === 'k12') {
        wsData = [
          [
            "STT", "Mã học sinh", "Họ và tên ", "Học kì", 
            "Năng lực chung", "", "", "", 
            "Năng lực đặc thù", "", "", "", "", "", 
            "Phẩm chất", "", "", "", "", ""
          ],
          [
            "", "", "", "", 
            "Nhận xét về năng lực", "Tự chủ và tự học", "Giao tiếp và hợp tác", "GQVĐ và sáng tạo", 
            "Nhận xét về năng lực đặc thù", "Ngôn ngữ", "Tính toán", "Khoa học", "Thẩm mĩ", "Thể chất", 
            "Nhận xét về phẩm chất", "Yêu nước", "Nhân ái", "Chăm chỉ", "Trung thực", "Trách nhiệm"
          ]
        ];

        students.forEach((std, idx) => {
          wsData.push([
            idx + 1,
            std.studentId || "",
            std.name,
            "Học kỳ 1", 
            std.comments.chung,
            std.individualComments["Tự chủ và tự học"] || "",
            std.individualComments["Giao tiếp và hợp tác"] || "",
            std.individualComments["Giải quyết vấn đề"] || "",
            std.comments.dacThu,
            std.individualComments["Ngôn ngữ"] || "",
            std.individualComments["Tính toán"] || "",
            std.individualComments["Khoa học"] || "",
            std.individualComments["Thẩm mĩ"] || "",
            std.individualComments["Thể chất"] || "",
            std.comments.phamChat,
            std.individualComments["Yêu nước"] || "",
            std.individualComments["Nhân ái"] || "",
            std.individualComments["Chăm chỉ"] || "",
            std.individualComments["Trung thực"] || "",
            std.individualComments["Trách nhiệm"] || ""
          ]);
        });

        merges = [
          { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },  { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, 
          { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },  { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, 
          { s: { r: 0, c: 4 }, e: { r: 0, c: 7 } }, 
          { s: { r: 0, c: 8 }, e: { r: 0, c: 13 } },
          { s: { r: 0, c: 14 }, e: { r: 0, c: 19 } }
        ];

        cols = [
          { wch: 5 }, { wch: 15 }, { wch: 22 }, { wch: 10 }, 
          { wch: 50 }, { wch: 25 }, { wch: 25 }, { wch: 25 },
          { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, 
          { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }
        ];

      } else {
        wsData = [
          [
            "STT", "Mã học sinh", "Họ và tên ", "Học kì", 
            "Năng lực chung", "", "", "", 
            "Năng lực đặc thù", "", "", "", "", "", "", "", 
            "Phẩm chất", "", "", "", "", ""
          ],
          [
            "", "", "", "", 
            "Nhận xét về năng lực", "Tự chủ và tự học", "Giao tiếp và hợp tác", "GQVĐ và sáng tạo", 
            "Nhận xét về năng lực đặc thù", "Ngôn ngữ", "Tính toán", "Khoa học", "Thẩm mĩ", "Thể chất", "Công nghệ", "Tin học",
            "Nhận xét về phẩm chất", "Yêu nước", "Nhân ái", "Chăm chỉ", "Trung thực", "Trách nhiệm"
          ]
        ];

        students.forEach((std, idx) => {
          wsData.push([
            idx + 1,
            std.studentId || "",
            std.name,
            "Học kỳ 1", 
            std.comments.chung,
            std.individualComments["Tự chủ và tự học"] || "",
            std.individualComments["Giao tiếp và hợp tác"] || "",
            std.individualComments["Giải quyết vấn đề"] || "",
            std.comments.dacThu,
            std.individualComments["Ngôn ngữ"] || "",
            std.individualComments["Tính toán"] || "",
            std.individualComments["Khoa học"] || "",
            std.individualComments["Thẩm mĩ"] || "",
            std.individualComments["Thể chất"] || "",
            std.individualComments["Công nghệ"] || "",
            std.individualComments["Tin học"] || "",
            std.comments.phamChat,
            std.individualComments["Yêu nước"] || "",
            std.individualComments["Nhân ái"] || "",
            std.individualComments["Chăm chỉ"] || "",
            std.individualComments["Trung thực"] || "",
            std.individualComments["Trách nhiệm"] || ""
          ]);
        });

        merges = [
          { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },  { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, 
          { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },  { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, 
          { s: { r: 0, c: 4 }, e: { r: 0, c: 7 } }, 
          { s: { r: 0, c: 8 }, e: { r: 0, c: 15 } }, 
          { s: { r: 0, c: 16 }, e: { r: 0, c: 21 } } 
        ];

        cols = [
          { wch: 5 }, { wch: 15 }, { wch: 22 }, { wch: 10 }, 
          { wch: 50 }, { wch: 25 }, { wch: 25 }, { wch: 25 },
          { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, 
          { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }
        ];
      }

      const worksheet = XLSX.utils.aoa_to_sheet(wsData);
      worksheet['!merges'] = merges;
      worksheet['!cols'] = cols;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tong_Hop");
      
      const outFileName = gradeBlock === 'k12' ? "Ket_Qua_Nhan_Xet_Khoi_1_2.xlsx" : "Ket_Qua_Nhan_Xet_Khoi_3_4_5.xlsx";
      XLSX.writeFile(workbook, outFileName);
    } catch (error: any) {
      alert("Lỗi xuất file: " + error.message);
    }
  };

  return (
    <div className="flex flex-col flex-grow bg-slate-50 font-sans text-slate-800" style={{ fontFamily: "'Inter', 'Roboto', sans-serif" }}>
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden relative flex flex-col flex-grow">
        
        {/* Header (Secondary) */}
        <div className="bg-indigo-600 p-6 text-white relative overflow-hidden">
          <div className="relative z-10">
            <h2 className="text-xl font-bold flex items-center gap-2">
              <CheckCircle2 className="w-6 h-6 text-indigo-200" />
              Nhận xét Năng lực & Phẩm chất
            </h2>
            <p className="mt-1 text-indigo-100 text-xs">
              Tự động hóa lời phê năng lực chung, năng lực đặc thù và phẩm chất từ bảng điểm Excel.
            </p>
          </div>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-slate-200 bg-slate-50">
          <button 
            className={`flex items-center gap-2 px-6 py-3 font-medium text-sm transition-colors ${activeTab === 'generator' ? 'bg-white text-indigo-700 border-b-2 border-indigo-700' : 'text-slate-500 hover:text-slate-700'}`}
            onClick={() => setActiveTab('generator')}
          >
            <Users className="w-4 h-4" />
            Tạo Nhận Xét & Xuất File
          </button>
          <button 
            className={`flex items-center gap-2 px-6 py-3 font-medium text-sm transition-colors ${activeTab === 'settings' ? 'bg-white text-indigo-700 border-b-2 border-indigo-700' : 'text-slate-500 hover:text-slate-700'}`}
            onClick={() => setActiveTab('settings')}
          >
            <Settings className="w-4 h-4" />
            Cấu hình Ngân hàng câu
          </button>
        </div>

        {/* Content */}
        <div className="p-6 flex flex-col flex-grow overflow-auto">
          
          {/* TAB 1: GENERATOR */}
          {activeTab === 'generator' && (
            <div className="space-y-6">
              
              <div className="bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-100 rounded-lg p-5 flex flex-col items-start">
                <h3 className="font-semibold text-blue-800 mb-4 flex items-center gap-2">
                  <Info className="w-5 h-5" /> Bước 1: Chọn Khối lớp & Tải lên bảng điểm
                </h3>
                
                {/* LỰA CHỌN KHỐI LỚP */}
                <div className="bg-white p-4 rounded-md border border-slate-200 w-full mb-5 flex flex-col sm:flex-row gap-6 shadow-sm">
                  <label className={`flex items-center gap-3 cursor-pointer p-2 rounded ${gradeBlock === 'k12' ? 'bg-blue-50 font-semibold text-blue-700' : 'text-slate-600 hover:bg-slate-50'}`}>
                    <input 
                      type="radio" 
                      name="gradeBlock" 
                      value="k12" 
                      checked={gradeBlock === 'k12'} 
                      onChange={() => setGradeBlock('k12')}
                      className="w-5 h-5 text-blue-600 focus:ring-blue-500"
                    />
                    Khối 1, 2 (Mẫu 20 Cột)
                  </label>
                  <label className={`flex items-center gap-3 cursor-pointer p-2 rounded ${gradeBlock === 'k345' ? 'bg-blue-50 font-semibold text-blue-700' : 'text-slate-600 hover:bg-slate-50'}`}>
                    <input 
                      type="radio" 
                      name="gradeBlock" 
                      value="k345" 
                      checked={gradeBlock === 'k345'} 
                      onChange={() => setGradeBlock('k345')}
                      className="w-5 h-5 text-blue-600 focus:ring-blue-500"
                    />
                    Khối 3, 4, 5 (Có Công nghệ, Tin học)
                  </label>
                </div>

                <div className="flex items-start gap-2 text-sm text-indigo-700 mb-4 bg-white p-3 rounded shadow-sm border border-indigo-100 w-full">
                  <Info className="w-5 h-5 text-yellow-500 flex-shrink-0" />
                  <p><b>Lưu ý:</b> Phần mềm sẽ tự động chọn câu nhận xét khác nhau cho cùng một mức đánh giá để tránh trùng lặp 100%.</p>
                </div>
                
                <input 
                  type="file" 
                  accept=".xlsx, .xls, .csv" 
                  ref={fileInputRef}
                  className="hidden"
                  onChange={handleFileUpload}
                />
                <button 
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-md font-medium text-sm flex items-center gap-2 transition-colors shadow-sm"
                >
                  <Upload className="w-5 h-5" />
                  Chọn File Bảng Điểm Tải Lên
                </button>
                
                {statusMsg.text && (
                  <div className={`mt-4 p-3 rounded-md text-sm font-medium flex items-center gap-2 w-full ${statusMsg.type === 'success' ? 'bg-green-100 text-green-700 border border-green-200' : statusMsg.type === 'info' ? 'bg-blue-100 text-blue-700 border border-blue-200' : 'bg-red-100 text-red-700 border border-red-200'}`}>
                    {statusMsg.type === 'success' ? <CheckCircle2 className="w-4 h-4" /> : <AlertCircle className="w-4 h-4" />}
                    {statusMsg.text}
                  </div>
                )}
              </div>

              {students.length > 0 && (
                <div className="border border-slate-300 rounded-lg overflow-hidden bg-white shadow-sm flex flex-col">
                  <div className="bg-slate-50 px-4 py-4 border-b border-slate-300 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                    <h3 className="font-semibold text-slate-800 text-sm">
                      Bước 2: Kết quả xem trước ({students.length} học sinh)
                    </h3>
                    <button 
                      onClick={exportToStandardTemplateExcel}
                      className="bg-green-600 hover:bg-green-700 text-white px-5 py-2.5 rounded-md text-xs font-medium flex items-center gap-2 transition-colors shadow-md"
                    >
                      <FileDown className="w-4 h-4" />
                      Xuất File Excel Kết Quả
                    </button>
                  </div>
                  
                  {/* BẢNG HIỂN THỊ */}
                  <div className="overflow-x-auto max-h-[500px]">
                    <table className="w-max min-w-full text-[11px] text-left text-slate-700 border-collapse">
                      <thead className="bg-slate-100 sticky top-0 z-10 shadow-sm text-center font-bold">
                        <tr>
                          <th rowSpan={2} className="border border-slate-300 px-2 py-2 w-10 bg-slate-200">STT</th>
                          <th rowSpan={2} className="border border-slate-300 px-2 py-2 w-14 bg-slate-200">Sửa</th>
                          <th rowSpan={2} className="border border-slate-300 px-3 py-2 w-40 bg-slate-200">Họ và tên</th>
                          
                          <th colSpan={4} className="border border-slate-300 px-2 py-2 bg-blue-100">Năng lực chung</th>
                          <th colSpan={gradeBlock === 'k12' ? 6 : 8} className="border border-slate-300 px-2 py-2 bg-green-100">Năng lực đặc thù</th>
                          <th colSpan={6} className="border border-slate-300 px-2 py-2 bg-amber-100">Phẩm chất</th>
                        </tr>
                        <tr className="text-[10px]">
                          <th className="border border-slate-300 px-2 py-1 bg-blue-50 w-48">Nhận xét gộp</th>
                          <th className="border border-slate-300 px-2 py-1 bg-blue-50 w-32">Tự chủ</th>
                          <th className="border border-slate-300 px-2 py-1 bg-blue-50 w-32">Giao tiếp</th>
                          <th className="border border-slate-300 px-2 py-1 bg-blue-50 w-32">Sáng tạo</th>
                          
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-48">Nhận xét gộp</th>
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Ngôn ngữ</th>
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Tính toán</th>
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Khoa học</th>
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Thẩm mĩ</th>
                          <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Thể chất</th>
                          {gradeBlock === 'k345' && <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Công nghệ</th>}
                          {gradeBlock === 'k345' && <th className="border border-slate-300 px-2 py-1 bg-green-50 w-32">Tin học</th>}
                          
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-48">Nhận xét gộp</th>
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-32">Yêu nước</th>
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-32">Nhân ái</th>
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-32">Chăm chỉ</th>
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-32">Trung thực</th>
                          <th className="border border-slate-300 px-2 py-1 bg-amber-50 w-32">Trách nhiệm</th>
                        </tr>
                      </thead>
                      <tbody>
                        {students.map((std, index) => (
                          <tr key={std.id} className="hover:bg-slate-50 transition-colors">
                            <td className="border border-slate-200 px-2 py-2 text-center bg-white">{index + 1}</td>
                            <td className="border border-slate-200 px-1 py-1 text-center bg-white">
                              <button onClick={() => openEditModal(std)} className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded">
                                <Pencil className="w-4 h-4" />
                              </button>
                            </td>
                            <td className="border border-slate-200 px-3 py-2 font-medium bg-white whitespace-nowrap">{std.name}</td>
                            
                            {/* Cột Chung */}
                            <td className="border border-slate-200 px-2 py-1 bg-blue-50/10 italic text-[10px]">{std.comments.chung}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Tự chủ và tự học"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Giao tiếp và hợp tác"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Giải quyết vấn đề"]}</td>
                            
                            {/* Cột Đặc thù */}
                            <td className="border border-slate-200 px-2 py-1 bg-green-50/10 italic text-[10px]">{std.comments.dacThu}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Ngôn ngữ"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Tính toán"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Khoa học"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Thẩm mĩ"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Thể chất"]}</td>
                            {gradeBlock === 'k345' && <td className="border border-slate-200 px-2 py-1">{std.individualComments["Công nghệ"]}</td>}
                            {gradeBlock === 'k345' && <td className="border border-slate-200 px-2 py-1">{std.individualComments["Tin học"]}</td>}
                            
                            {/* Cột Phẩm chất */}
                            <td className="border border-slate-200 px-2 py-1 bg-amber-50/10 italic text-[10px]">{std.comments.phamChat}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Yêu nước"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Nhân ái"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Chăm chỉ"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Trung thực"]}</td>
                            <td className="border border-slate-200 px-2 py-1">{std.individualComments["Trách nhiệm"]}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB 2: SETTINGS */}
          {activeTab === 'settings' && (
            <div className="space-y-4 flex flex-col flex-grow">
              <div className="flex items-start gap-3 bg-yellow-50 text-yellow-800 p-4 rounded-lg text-sm border border-yellow-200">
                <AlertCircle className="w-5 h-5 flex-shrink-0 mt-0.5" />
                <div>
                  <p className="font-semibold text-base mb-1">Cấu hình ngân hàng câu hỏi</p>
                  <p>Mỗi tiêu chí nên có nhiều mẫu câu. Phần mềm sẽ chọn ngẫu nhiên để đảm bảo tính cá nhân hóa.</p>
                </div>
              </div>
              
              <textarea
                value={commentBankText}
                onChange={(e) => setCommentBankText(e.target.value)}
                className="w-full flex-grow p-4 border border-slate-300 rounded-lg font-mono text-xs focus:ring-2 focus:ring-indigo-500 min-h-[400px]"
                spellCheck="false"
              />
            </div>
          )}

        </div>
      </div>

      {/* MODAL EDIT */}
      {editingStudent && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-indigo-600 text-white">
              <h3 className="font-bold">Sửa nhận xét: {editingStudent.name}</h3>
              <button onClick={closeEditModal} className="text-white hover:opacity-75"><X /></button>
            </div>
            <div className="p-6 overflow-y-auto space-y-6 bg-slate-50 flex-grow">
              
              {/* Group Gộp */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="bg-blue-50 p-3 rounded border border-blue-200">
                  <span className="block text-[10px] font-bold text-blue-800 mb-2">NHẬN XÉT NĂNG LỰC CHUNG</span>
                  <textarea value={editingStudent.comments.chung} onChange={(e) => handleEditGroupedComment('chung', e.target.value)} className="w-full text-xs p-2 border rounded h-20" />
                </div>
                <div className="bg-green-50 p-3 rounded border border-green-200">
                  <span className="block text-[10px] font-bold text-green-800 mb-2">NHẬN XÉT NĂNG LỰC ĐẶC THÙ</span>
                  <textarea value={editingStudent.comments.dacThu} onChange={(e) => handleEditGroupedComment('dacThu', e.target.value)} className="w-full text-xs p-2 border rounded h-20" />
                </div>
                <div className="bg-amber-50 p-3 rounded border border-amber-200">
                  <span className="block text-[10px] font-bold text-amber-800 mb-2">NHẬN XÉT PHẨM CHẤT</span>
                  <textarea value={editingStudent.comments.phamChat} onChange={(e) => handleEditGroupedComment('phamChat', e.target.value)} className="w-full text-xs p-2 border rounded h-20" />
                </div>
              </div>

              <div className="flex justify-start">
                <button onClick={regenerateGroupedForEditingStudent} className="text-xs text-indigo-600 font-medium underline">Đồng bộ lại từ các tiêu chí riêng lẻ</button>
              </div>

              {/* Individual */}
              <div className="space-y-4">
                <span className="block font-bold text-sm border-b pb-1">Chi tiết từng tiêu chí</span>
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                  {Object.keys(editingStudent.individualComments).map(crit => (
                     <div key={crit} className="space-y-1">
                        <span className="text-[9px] font-bold text-slate-500 uppercase">{crit}</span>
                        <textarea 
                          value={editingStudent.individualComments[crit]} 
                          onChange={(e) => handleEditIndividualComment(crit, e.target.value)}
                          className="w-full text-[10px] p-2 border rounded h-14"
                        />
                     </div>
                  ))}
                </div>
              </div>

            </div>
            <div className="px-6 py-4 border-t bg-white flex justify-end gap-3">
              <button onClick={closeEditModal} className="px-4 py-2 text-sm text-slate-500">Hủy</button>
              <button onClick={saveEditedStudent} className="px-6 py-2 bg-indigo-600 text-white rounded font-medium flex items-center gap-2">
                <Save className="w-4 h-4" /> Lưu thay đổi
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
