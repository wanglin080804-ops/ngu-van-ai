import React, { useState, useRef, useEffect } from 'react';
import { marked } from 'marked';
import katex from 'katex';
import DOMPurify from 'dompurify';
import PptxGenJS from 'pptxgenjs';
import { 
  FileText, 
  Sparkles, 
  BookOpen, 
  Download, 
  Copy, 
  Check, 
  Printer, 
  X,
  BookMarked
} from 'lucide-react';

interface HistoryItem {
  id: string;
  timestamp: number;
  taskType: 'khbd' | 'matran' | 'slide';
  bookFile: string;
  lessonName: string;
  content: string;
}

const BOOKS = [
  { label: 'Ngá»¯ vÄƒn 6 Táº­p 1 (Káº¿t ná»‘i tri thá»©c)', value: 'SGK NV6 T1.pdf', grade: 'Lá»›p 6' },
  { label: 'Ngá»¯ vÄƒn 6 Táº­p 2 (Káº¿t ná»‘i tri thá»©c)', value: 'SGK NV6 T2.pdf', grade: 'Lá»›p 6' },
  { label: 'Ngá»¯ vÄƒn 9 Táº­p 1 (Káº¿t ná»‘i tri thá»©c)', value: 'SGK NV9 T1.pdf', grade: 'Lá»›p 9' },
  { label: 'Ngá»¯ vÄƒn 9 Táº­p 2 (Káº¿t ná»‘i tri thá»©c)', value: 'SGK NV9 T2.pdf', grade: 'Lá»›p 9' }
];

export default function App() {
  const [taskType, setTaskType] = useState<'khbd' | 'matran' | 'slide'>('khbd');
  const [selectedBook, setSelectedBook] = useState<string>(BOOKS[0].value);
  const [lessonName, setLessonName] = useState<string>('');
  const [extraContext, setExtraContext] = useState<string>('');
  const [pageRange, setPageRange] = useState<string>('');
  
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [resultContent, setResultContent] = useState<string>('');
  const [copied, setCopied] = useState<boolean>(false);
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);

  const resultContainerRef = useRef<HTMLDivElement>(null);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => {
      setNotification(null);
    }, 3500);
  };

  const renderFormattedContent = (markdownText: string) => {
    if (!markdownText) return '';

    const mathTokens: string[] = [];
    
    let textWithMath = markdownText.replace(/\$\$([\s\S]+?)\$\$/g, (_match, formula) => {
      try {
        const html = katex.renderToString(formula.trim(), { displayMode: true, throwOnError: false });
        mathTokens.push(html);
        return `MATHTOKENPLACEHOLDER${mathTokens.length - 1}END`;
      } catch {
        return _match;
      }
    });

    textWithMath = textWithMath.replace(/\$([^\$\n\r]+?)\$/g, (_match, formula) => {
      try {
        const html = katex.renderToString(formula.trim(), { displayMode: false, throwOnError: false });
        mathTokens.push(html);
        return `MATHTOKENPLACEHOLDER${mathTokens.length - 1}END`;
      } catch {
        return _match;
      }
    });

    let html = '';
    try {
      html = marked.parse(textWithMath, { async: false, breaks: true }) as string;
    } catch {
      html = marked.parse(markdownText, { async: false, breaks: true }) as string;
    }

    mathTokens.forEach((tokenHtml, index) => {
      html = html.replace(`MATHTOKENPLACEHOLDER${index}END`, tokenHtml);
    });

    return html;
  };

  const generateSlidePPTX = async (jsonText: string) => {
    try {
      const slidesData = JSON.parse(jsonText);
      if (!Array.isArray(slidesData)) throw new Error("Format JSON khÃ´ng há»£p lá»‡");

      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';

      slidesData.forEach((slide) => {
        const pptSlide = pptx.addSlide();
        
        // Title
        pptSlide.addText(slide.title || '', {
          x: 0.5, y: 0.5, w: '90%', h: 1,
          fontSize: 32, bold: true, color: '003366', align: 'center'
        });

        // Content
        pptSlide.addText(slide.content || '', {
          x: 0.5, y: 1.8, w: '90%', h: 4,
          fontSize: 20, color: '333333', align: 'left', bullet: true, valign: 'top'
        });

        // Notes
        if (slide.notes) {
          pptSlide.addNotes(slide.notes);
        }
      });

      const fileName = `Slide_${lessonName.replace(/[\/\\?%*:|"<>]/g, '_')}.pptx`;
      await pptx.writeFile({ fileName });
      showToast("Táº£i xuá»‘ng PowerPoint thÃ nh cÃ´ng!", "success");
    } catch (err) {
      console.error(err);
      showToast("Lá»—i khi táº¡o file PPTX tá»« dá»¯ liá»‡u: " + err, "error");
      setResultContent("Dá»¯ liá»‡u JSON sinh ra khÃ´ng há»£p lá»‡:\n\n" + jsonText);
    }
  };

  const handleGenerate = async () => {
    if (!lessonName.trim()) {
      showToast("Vui lÃ²ng Ä‘iá»n tÃªn bÃ i há»c!", "error");
      return;
    }

    const book = BOOKS.find(b => b.value === selectedBook);
    if (!book) return;

    setIsGenerating(true);
    setResultContent('');

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          taskType,
          subject: 'Ngá»¯ vÄƒn',
          grade: book.grade,
          lessonName: lessonName.trim(),
          extraContext: extraContext.trim(),
          selectedBookFile: selectedBook,
          pageRange: pageRange.trim()
        })
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || "KhÃ´ng thá»ƒ táº¡o tÃ i liá»‡u. Vui lÃ²ng thá»­ láº¡i!");
      }

      let generatedText = data.content || '';
      
      // Clean markdown JSON block if necessary
      if (taskType === 'slide') {
         generatedText = generatedText.replace(/^```json/, '').replace(/```$/, '').trim();
      }

      setResultContent(generatedText);

      if (taskType === 'slide') {
        await generateSlidePPTX(generatedText);
      } else {
        showToast("ðŸŽ‰ ÄÃ£ táº¡o tÃ i liá»‡u thÃ nh cÃ´ng!", "success");
        setTimeout(() => {
          resultContainerRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 150);
      }

    } catch (err: any) {
      showToast(err.message || "ÄÃ£ xáº£y ra lá»—i khi táº¡o tÃ i liá»‡u.", "error");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleCopy = () => {
    if (!resultContent) return;
    navigator.clipboard.writeText(resultContent).then(() => {
      setCopied(true);
      showToast("âœ… ÄÃ£ copy toÃ n bá»™ ná»™i dung!", "success");
      setTimeout(() => setCopied(false), 2500);
    }).catch(err => {
      showToast("Lá»—i khi copy: " + err, "error");
    });
  };

  const handleExportWord = () => {
    if (!resultContent) return;
    const renderedHtml = renderFormattedContent(resultContent);
    const title = taskType === 'khbd' 
      ? `KHBD_${lessonName}`
      : `MaTran_${lessonName}`;
    
    const preHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${title}</title>
        <style>
          body { font-family: 'Times New Roman', serif; font-size: 13pt; line-height: 1.5; color: #000; margin: 2cm; }
          h1 { font-size: 16pt; text-align: center; color: #002060; font-weight: bold; text-transform: uppercase; margin-bottom: 12pt; }
          h2 { font-size: 14pt; color: #002060; font-weight: bold; margin-top: 14pt; margin-bottom: 6pt; }
          h3 { font-size: 13pt; font-weight: bold; margin-top: 10pt; margin-bottom: 4pt; }
          table { border-collapse: collapse; width: 100%; margin-top: 10pt; margin-bottom: 12pt; }
          th, td { border: 1px solid black; padding: 6pt 8pt; vertical-align: top; font-size: 11pt; }
          th { background-color: #F2F2F2; font-weight: bold; text-align: center; }
          p { margin-bottom: 6pt; margin-top: 0; }
          ul, ol { margin-top: 0; margin-bottom: 6pt; padding-left: 20pt; }
          blockquote { border-left: 3px solid #666; padding-left: 10pt; margin-left: 0; color: #444; font-style: italic; }
        </style>
      </head>
      <body>`;
    const postHtml = `</body></html>`;
    const fullHtml = preHtml + renderedHtml + postHtml;

    const blob = new Blob(['\ufeff' + fullHtml], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.download = `${title.replace(/[\/\\?%*:|"<>]/g, '_')}.doc`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
    showToast("ðŸ“¥ ÄÃ£ táº£i file Word (.doc) thÃ nh cÃ´ng!", "success");
  };

  const handlePrintOrPdf = () => {
    window.print();
  };

  return (
    <div className="min-h-screen py-10 px-4 sm:px-6 flex justify-center items-start">
      {notification && (
        <div className={`fixed top-6 right-6 z-50 px-6 py-4 rounded-xl shadow-2xl flex items-center gap-3 text-sm font-medium transition-all animate-bounce ${
          notification.type === 'error' ? 'bg-red-500 text-white' :
          notification.type === 'info' ? 'bg-indigo-500 text-white' : 'bg-emerald-500 text-white'
        }`}>
          <span>{notification.message}</span>
        </div>
      )}

      <div className="max-w-4xl w-full">
        {/* Main Header */}
        <div className="text-center mb-10 mt-6 float-anim">
          <h1 className="text-4xl sm:text-5xl font-extrabold tracking-tight mb-4 bg-gradient-to-r from-purple-400 via-pink-400 to-indigo-400 bg-clip-text text-transparent">
            Ngá»¯ VÄƒn Tháº§y Ãšt - Káº¿t ná»‘i tri thá»©c
          </h1>
          <p className="text-slate-300 text-base sm:text-lg font-light max-w-2xl mx-auto">
            Há»‡ thá»‘ng há»— trá»£ táº¡o Káº¿ hoáº¡ch bÃ i dáº¡y, Slide bÃ i giáº£ng (PowerPoint) vÃ  Ma tráº­n Ä‘á» kiá»ƒm tra Ä‘á»™c quyá»n cho mÃ´n Ngá»¯ VÄƒn lá»›p 6 & 9.
          </p>
        </div>

        {/* Input Panel */}
        <div className="glass-panel p-8 mb-10 no-print pulse-glow">
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
            <div>
              <label className="block text-sm font-semibold text-slate-300 mb-2 flex items-center gap-2">
                <BookOpen className="w-4 h-4 text-purple-400" />
                Chá»n SÃ¡ch GiÃ¡o Khoa
              </label>
              <select
                value={selectedBook}
                onChange={(e) => setSelectedBook(e.target.value)}
                className="w-full px-4 py-3 rounded-xl input-glass font-medium text-sm"
              >
                {BOOKS.map(b => (
                  <option key={b.value} value={b.value}>{b.label}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-300 mb-2 flex items-center gap-2">
                <FileText className="w-4 h-4 text-purple-400" />
                Loáº¡i tÃ i liá»‡u cáº§n táº¡o
              </label>
              <select
                value={taskType}
                onChange={(e) => setTaskType(e.target.value as 'khbd' | 'matran' | 'slide')}
                className="w-full px-4 py-3 rounded-xl input-glass font-medium text-sm"
              >
                <option value="khbd">Káº¿ hoáº¡ch bÃ i dáº¡y (CÃ´ng vÄƒn 5512)</option>
                <option value="slide">Slide BÃ i Giáº£ng (Táº£i file PowerPoint trá»±c tiáº¿p)</option>
                <option value="matran">Ma tráº­n & Äáº·c táº£ Ä‘á» kiá»ƒm tra</option>
              </select>
            </div>
          </div>

          <div className="mb-6">
            <label className="block text-sm font-semibold text-slate-300 mb-2 flex items-center gap-2">
              <BookMarked className="w-4 h-4 text-purple-400" />
              TÃªn bÃ i há»c / Chá»§ Ä‘á»
            </label>
            <input
              type="text"
              value={lessonName}
              onChange={(e) => setLessonName(e.target.value)}
              placeholder="VD: BÃ i 1: TÃ´i vÃ  cÃ¡c báº¡n - Truyá»‡n Ä‘á»“ng thoáº¡i..."
              className="w-full px-4 py-3 rounded-xl input-glass text-sm"
            />
          </div>

          <div className="mb-6">
            <label className="block text-sm font-semibold text-slate-300 mb-2">
              Khoáº£ng trang (TÃ¹y chá»n)
            </label>
            <input
              type="text"
              value={pageRange}
              onChange={(e) => setPageRange(e.target.value)}
              placeholder="VD: Tá»« trang 15 Ä‘áº¿n trang 18"
              className="w-full px-4 py-3 rounded-xl input-glass text-sm"
            />
            <p className="text-xs text-slate-400 mt-2">Giá»›i háº¡n sá»‘ trang giÃºp AI Ä‘á»c ná»™i dung SGK chÃ­nh xÃ¡c vÃ  nhanh chÃ³ng hÆ¡n.</p>
          </div>

          <div className="mb-8">
            <label className="block text-sm font-semibold text-slate-300 mb-2">
              YÃªu cáº§u bá»• sung (TÃ¹y chá»n)
            </label>
            <textarea
              value={extraContext}
              onChange={(e) => setExtraContext(e.target.value)}
              placeholder="Ghi chÃº thÃªm vá» yÃªu cáº§u cho tÃ i liá»‡u..."
              className="w-full px-4 py-3 rounded-xl input-glass text-sm h-24 resize-none"
            />
          </div>

          <button
            onClick={handleGenerate}
            disabled={isGenerating}
            className={`w-full py-4 rounded-xl text-base font-bold transition-all shadow-lg flex items-center justify-center gap-3 ${
              isGenerating ? 'bg-slate-700 text-slate-300 cursor-not-allowed' : 'btn-primary cursor-pointer'
            }`}
          >
            {isGenerating ? (
              <>
                <div className="loader-spinner"></div>
                Äang phÃ¢n tÃ­ch sÃ¡ch vÃ  táº¡o tÃ i liá»‡u...
              </>
            ) : (
              <>
                <Sparkles className="w-5 h-5" />
                {taskType === 'slide' ? 'Táº¡o & Táº£i Xuá»‘ng Slide PowerPoint' : 'Báº¯t Äáº§u Táº¡o TÃ i Liá»‡u AI'}
              </>
            )}
          </button>
        </div>

        {/* Result Area */}
        {resultContent && taskType !== 'slide' && (
          <div ref={resultContainerRef} className="glass-panel overflow-hidden mb-12">
            <div className="bg-slate-900/50 px-6 py-10 flex flex-col items-center justify-center no-print">
              <div className="w-16 h-16 bg-emerald-500/20 rounded-full flex items-center justify-center mb-5">
                <Check className="w-8 h-8 text-emerald-400" />
              </div>
              <h3 className="font-bold text-2xl text-white mb-3">
                Táº¡o tÃ i liá»‡u thÃ nh cÃ´ng!
              </h3>
              <p className="text-slate-300 mb-8 text-center max-w-md text-base">
                TÃ i liá»‡u cá»§a tháº§y Ä‘Ã£ sáºµn sÃ ng. Vui lÃ²ng chá»n Ä‘á»‹nh dáº¡ng muá»‘n táº£i xuá»‘ng.
              </p>
              <div className="flex flex-col sm:flex-row gap-4">
                <button
                  onClick={handleExportWord}
                  className="px-6 py-3.5 bg-gradient-to-r from-indigo-500 to-purple-600 hover:from-indigo-600 hover:to-purple-700 text-white rounded-xl font-bold transition-all shadow-lg flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Download className="w-5 h-5" />
                  Táº£i file Word (.doc)
                </button>
                <button
                  onClick={handlePrintOrPdf}
                  className="px-6 py-3.5 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-xl font-bold transition-all flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Printer className="w-5 h-5" />
                  LÆ°u file PDF
                </button>
              </div>
            </div>
            
            {/* Hidden on screen, visible when printing */}
            <div className="hidden print:block p-8 bg-white text-slate-900">
              <div 
                className="markdown-body"
                dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(renderFormattedContent(resultContent)) }} 
              />
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="text-center pb-8 pt-4 text-slate-400/80 text-sm no-print">
          <p className="font-semibold text-slate-300 mb-1">Tháº§y Tráº§m Thanh Ãšt</p>
          <p>ÄVCT: TrÆ°á»ng THCS PhÆ°á»›c HÆ°ng</p>
        </div>
      </div>
    </div>
  );
}

