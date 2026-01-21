import React from 'react';
import { ProcessingStatus } from '../types';

interface StatusCardProps {
  status: ProcessingStatus;
  progress: number;
  fileName: string | null;
  error?: string;
}

export const StatusCard: React.FC<StatusCardProps> = ({ status, progress, fileName, error }) => {
  if (status === ProcessingStatus.IDLE) return null;

  const isError = status === ProcessingStatus.ERROR;
  const isComplete = status === ProcessingStatus.COMPLETED;

  return (
    <div className={`mt-8 p-6 rounded-xl border shadow-sm ${isError ? 'bg-red-50 border-red-200' : 'bg-white border-slate-200'}`}>
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center space-x-3">
          {isComplete ? (
             <div className="bg-green-100 p-2 rounded-full">
               <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                 <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
               </svg>
             </div>
          ) : isError ? (
            <div className="bg-red-100 p-2 rounded-full">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-red-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
              </svg>
            </div>
          ) : (
            <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
          )}
          <div>
            <h3 className="font-semibold text-slate-800">
              {isError ? "Conversion Failed" : isComplete ? "Conversion Successful" : "Processing..."}
            </h3>
            <p className="text-sm text-slate-500">{fileName}</p>
          </div>
        </div>
        <span className="text-sm font-medium text-slate-600">{Math.round(progress)}%</span>
      </div>

      {!isError && (
        <div className="w-full bg-slate-200 rounded-full h-2.5 mb-2 overflow-hidden">
          <div 
            className={`h-2.5 rounded-full transition-all duration-300 ${isComplete ? 'bg-green-500' : 'bg-blue-600'}`}
            style={{ width: `${progress}%` }}
          ></div>
        </div>
      )}

      <p className="text-sm text-slate-500 mt-2">
        {status === ProcessingStatus.READING_PDF && "Extracting high-fidelity page images..."}
        {status === ProcessingStatus.ANALYZING_AI && "Gemini AI is generating speaker notes..."}
        {status === ProcessingStatus.GENERATING_PPTX && "Assembling PowerPoint file..."}
        {status === ProcessingStatus.COMPLETED && "Download started automatically."}
        {status === ProcessingStatus.ERROR && error}
      </p>
    </div>
  );
};