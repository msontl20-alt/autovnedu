import { useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  GraduationCap, 
  CheckCircle, 
  BookOpen, 
  Star,
  Settings,
  Info
} from 'lucide-react';
import SubjectCommentApp from './SubjectCommentApp';
import CompetencyApp from './CompetencyApp';
import { cn } from '@/src/lib/utils';

type AppTab = 'subject' | 'competency';

export default function App() {
  const [activeTab, setActiveTab] = useState<AppTab>('subject');

  return (
    <div className="flex flex-col min-h-screen bg-gray-50">
      {/* Global Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-20">
            {/* Logo Section */}
            <div className="flex items-center gap-4">
              <div className="bg-blue-600 p-2.5 rounded-2xl shadow-lg shadow-blue-100 transform -rotate-3">
                <GraduationCap size={28} className="text-white" />
              </div>
              <div className="flex flex-col">
                <h1 className="text-xl font-black text-gray-900 tracking-tight leading-none">
                  SmartComment <span className="text-blue-600">Pro</span>
                </h1>
                <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-1">
                  Trợ lý nhận xét học sinh thông minh
                </p>
              </div>
            </div>

            {/* Tab Switcher */}
            <nav className="flex bg-gray-100/80 p-1.5 rounded-2xl border border-gray-200/50 backdrop-blur-md">
              <button
                onClick={() => setActiveTab('subject')}
                className={cn(
                  "flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-bold transition-all duration-300",
                  activeTab === 'subject'
                    ? "bg-white text-blue-600 shadow-sm ring-1 ring-black/5"
                    : "text-gray-500 hover:text-gray-700 hover:bg-gray-200/50"
                )}
              >
                <BookOpen size={18} />
                <span>Nhận xét Môn học</span>
              </button>
              <button
                onClick={() => setActiveTab('competency')}
                className={cn(
                  "flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-bold transition-all duration-300 ml-1",
                  activeTab === 'competency'
                    ? "bg-white text-indigo-600 shadow-sm ring-1 ring-black/5"
                    : "text-gray-500 hover:text-gray-700 hover:bg-gray-200/50"
                )}
              >
                <Star size={18} />
                <span>Năng lực & Phẩm chất</span>
              </button>
            </nav>

            {/* Action Group */}
            <div className="hidden md:flex items-center gap-4">
              <div className="flex items-center gap-2 bg-green-50 px-3 py-1.5 rounded-full border border-green-100 text-[11px] font-bold text-green-700">
                <CheckCircle size={14} className="text-green-500" />
                <span>Thông tư 27/2020</span>
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-grow flex flex-col p-4 md:p-6 lg:p-8 max-w-7xl mx-auto w-full">
        <AnimatePresence mode="wait">
          <motion.div
            key={activeTab}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            transition={{ duration: 0.2 }}
            className="flex-grow flex flex-col"
          >
            {activeTab === 'subject' ? <SubjectCommentApp /> : <CompetencyApp />}
          </motion.div>
        </AnimatePresence>
      </main>

      {/* Global Footer */}
      <footer className="bg-white border-t border-gray-100 py-6">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2 text-xs text-gray-400 font-medium">
              <Settings size={14} />
              <span>Phiên bản 2.5.0</span>
            </div>
            <div className="flex items-center gap-2 text-xs text-gray-400 font-medium">
              <Info size={14} />
              <span>Hỗ trợ VnEdu</span>
            </div>
          </div>
          <p className="text-[11px] text-gray-300 font-medium uppercase tracking-wider">
            © 2024 Design for Excellence in Education | Tác giả: Lục Minh Sơn
          </p>
        </div>
      </footer>

      {/* Global Styles for Animations */}
      <style>{`
        @keyframes pulse-soft {
          0%, 100% { transform: scale(1); opacity: 1; }
          50% { transform: scale(1.02); opacity: 0.9; }
        }
        .pulse-animation {
          animation: pulse-soft 3s infinite ease-in-out;
        }
        .table-container::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        .table-container::-webkit-scrollbar-track {
          background: #f1f1f1;
          border-radius: 10px;
        }
        .table-container::-webkit-scrollbar-thumb {
          background: #cbd5e1;
          border-radius: 10px;
        }
        .table-container::-webkit-scrollbar-thumb:hover {
          background: #94a3b8;
        }
      `}</style>
    </div>
  );
}
