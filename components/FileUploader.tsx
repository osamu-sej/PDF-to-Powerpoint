import React, { useCallback } from 'react';

interface FileUploaderProps {
  onFileSelect: (file: File) => void;
  disabled: boolean;
}

export const FileUploader: React.FC<FileUploaderProps> = ({ onFileSelect, disabled }) => {
  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      if (disabled) return;
      
      const file = e.dataTransfer.files[0];
      if (file && file.type === 'application/pdf') {
        onFileSelect(file);
      }
    },
    [disabled, onFileSelect]
  );

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (disabled) return;
    const file = e.target.files?.[0];
    if (file) {
      onFileSelect(file);
    }
  };

  return (
    <div
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      className={`
        border-2 border-dashed rounded-xl p-10 text-center transition-all cursor-pointer
        ${disabled 
          ? 'bg-slate-100 border-slate-300 opacity-50 cursor-not-allowed' 
          : 'bg-white border-blue-400 hover:border-blue-600 hover:bg-blue-50 shadow-sm'
        }
      `}
    >
      <input
        type="file"
        accept="application/pdf"
        onChange={handleInputChange}
        className="hidden"
        id="fileInput"
        disabled={disabled}
      />
      <label htmlFor="fileInput" className={`flex flex-col items-center justify-center ${disabled ? 'cursor-not-allowed' : 'cursor-pointer'}`}>
        <div className="bg-blue-100 p-4 rounded-full mb-4">
          <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
        </div>
        <p className="text-xl font-semibold text-slate-700 mb-2">
          Drop your PDF here
        </p>
        <p className="text-sm text-slate-500 mb-4">
          or click to browse files
        </p>
        <div className="text-xs text-blue-600 font-medium px-3 py-1 bg-blue-50 rounded-full">
          Supports .PDF only
        </div>
      </label>
    </div>
  );
};