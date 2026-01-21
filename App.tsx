import React, { useState } from 'react';
import { FileUploader } from './components/FileUploader';
import { StatusCard } from './components/StatusCard';
import { ApiKeyModal } from './components/ApiKeyModal';
import { ProcessingStatus, AppState, ProcessedPage } from './types';
import { convertPdfToImages } from './services/pdfService';
import { generateSlideNotes } from './services/geminiService';
import { generatePptxFile } from './services/pptxService';

const App: React.FC = () => {
  const [state, setState] = useState<AppState>({
    file: null,
    status: ProcessingStatus.IDLE,
    progress: 0,
    processedPages: [],
    useAI: false, 
    apiKey: '',
    showApiKeyModal: false
  });

  const handleToggleAI = () => {
    if (state.status !== ProcessingStatus.IDLE) return;

    if (state.useAI) {
      // Turning OFF
      setState(prev => ({ ...prev, useAI: false }));
    } else {
      // Turning ON
      if (state.apiKey) {
        // Key already exists, just enable
        setState(prev => ({ ...prev, useAI: true }));
      } else {
        // Need key, show modal
        setState(prev => ({ ...prev, showApiKeyModal: true }));
      }
    }
  };

  const handleSaveApiKey = (key: string) => {
    setState(prev => ({
      ...prev,
      apiKey: key,
      useAI: true,
      showApiKeyModal: false
    }));
  };

  const handleCancelApiKey = () => {
    setState(prev => ({
      ...prev,
      showApiKeyModal: false,
      useAI: false
    }));
  };

  const handleFileSelect = async (file: File) => {
    setState(prev => ({
      ...prev,
      file,
      status: ProcessingStatus.READING_PDF,
      progress: 0,
      errorMessage: undefined,
      processedPages: []
    }));

    try {
      // Step 1: Convert PDF to Images & Extract Text
      const pages = await convertPdfToImages(file, (pdfProgress) => {
        // PDF reading is the first 40% of the total progress
        setState(prev => ({ ...prev, progress: pdfProgress * 0.4 }));
      });

      let updatedPages = [...pages];

      // Step 2: AI Analysis (Optional)
      if (state.useAI) {
        setState(prev => ({ 
          ...prev, 
          status: ProcessingStatus.ANALYZING_AI,
          processedPages: pages
        }));
        
        // Process pages sequentially to strictly control rate limits
        const processAIForPage = async (page: ProcessedPage) => {
           // Pass the user's API key if available
           const notes = await generateSlideNotes(page.imageData, page.pageNumber, state.apiKey);
           return { ...page, aiNotes: notes };
        };

        const processedWithAI: ProcessedPage[] = [];
        
        // Sequential processing loop
        for (let i = 0; i < updatedPages.length; i++) {
            const page = updatedPages[i];
            const result = await processAIForPage(page);
            processedWithAI.push(result);
            
            // Add a small delay between requests to be gentle on the API
            if (i < updatedPages.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
            
            // AI is from 40% to 90% of total progress
            const aiProgress = ((processedWithAI.length / updatedPages.length) * 50) + 40;
            setState(prev => ({ ...prev, progress: aiProgress }));
        }
        updatedPages = processedWithAI;
      } else {
          // If no AI, jump progress
          setState(prev => ({ ...prev, progress: 90 }));
      }

      // Step 3: Generate PPTX
      setState(prev => ({ 
        ...prev, 
        status: ProcessingStatus.GENERATING_PPTX,
        processedPages: updatedPages
      }));

      await generatePptxFile(updatedPages, file.name);

      setState(prev => ({ 
        ...prev, 
        status: ProcessingStatus.COMPLETED, 
        progress: 100 
      }));

    } catch (err: any) {
      console.error(err);
      setState(prev => ({ 
        ...prev, 
        status: ProcessingStatus.ERROR, 
        errorMessage: err.message || "An unexpected error occurred." 
      }));
    }
  };

  const reset = () => {
    setState(prev => ({
      ...prev,
      file: null,
      status: ProcessingStatus.IDLE,
      progress: 0,
      processedPages: [],
      errorMessage: undefined
      // Keep useAI and apiKey setting for convenience
    }));
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 py-12 px-4 sm:px-6 lg:px-8 relative">
      {state.showApiKeyModal && (
        <ApiKeyModal onSave={handleSaveApiKey} onCancel={handleCancelApiKey} />
      )}

      <div className="max-w-3xl mx-auto">
        <div className="text-center mb-12">
          <h1 className="text-4xl font-extrabold text-slate-900 tracking-tight sm:text-5xl mb-4">
            PDF to PowerPoint <span className="text-blue-600">Pro</span>
          </h1>
          <p className="text-lg text-slate-600 max-w-2xl mx-auto">
            Convert your PDF slides into editable PowerPoint files with 
            <span className="font-semibold text-slate-800"> separated editable text</span> and faithful layout.
            <br />
            Now powered by Gemini AI to automatically generate speaker notes.
          </p>
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 p-8">
          
          {/* Controls */}
          <div 
             onClick={handleToggleAI}
             className={`mb-8 flex items-center justify-between bg-slate-50 p-4 rounded-lg border border-slate-200 transition-colors ${state.status === ProcessingStatus.IDLE ? 'cursor-pointer hover:bg-slate-100' : 'cursor-not-allowed opacity-75'}`}
          >
             <div className="flex items-center space-x-3">
               <div className={`p-2 rounded-lg transition-colors ${state.useAI ? 'bg-indigo-100 text-indigo-700' : 'bg-slate-200 text-slate-500'}`}>
                 <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                   <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-11a1 1 0 10-2 0v2H7a1 1 0 100 2h2v2a1 1 0 102 0v-2h2a1 1 0 100-2h-2V7z" clipRule="evenodd" />
                 </svg>
               </div>
               <div className="flex flex-col">
                 <span className="font-semibold text-slate-800">AI Enhancement</span>
                 <span className="text-xs text-slate-500">Generate speaker notes automatically</span>
               </div>
             </div>
             
             <div className="relative inline-flex items-center">
               <input 
                 type="checkbox" 
                 className="sr-only peer"
                 checked={state.useAI}
                 onChange={() => {}} // Handled by parent div onClick
                 disabled={state.status !== ProcessingStatus.IDLE}
               />
               <div className="w-11 h-6 bg-slate-200 peer-focus:outline-none peer-focus:ring-4 peer-focus:ring-blue-300 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
             </div>
          </div>

          {state.status === ProcessingStatus.IDLE || state.status === ProcessingStatus.COMPLETED || state.status === ProcessingStatus.ERROR ? (
            <div className="space-y-6">
               {state.status === ProcessingStatus.COMPLETED && (
                  <div className="text-center mb-6">
                    <button 
                      onClick={reset}
                      className="inline-flex items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-colors"
                    >
                      Convert Another File
                    </button>
                  </div>
               )}
               <FileUploader 
                onFileSelect={handleFileSelect} 
                disabled={false} 
              />
            </div>
          ) : (
            <div className="py-12">
              <div className="text-center">
                 <p className="text-slate-500 mb-4">Please wait while we process your document.</p>
              </div>
            </div>
          )}

          <StatusCard 
            status={state.status} 
            progress={state.progress} 
            fileName={state.file?.name || null}
            error={state.errorMessage}
          />
        </div>
        
        <div className="mt-8 text-center text-sm text-slate-400">
           <p>Privacy First: Files are processed locally in your browser (except when using AI features).</p>
        </div>
      </div>
    </div>
  );
};

export default App;