import React, { useState } from 'react';

interface ApiKeyModalProps {
  onSave: (key: string) => void;
  onCancel: () => void;
}

export const ApiKeyModal: React.FC<ApiKeyModalProps> = ({ onSave, onCancel }) => {
  const [key, setKey] = useState('');

  return (
    <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-6 animate-in fade-in zoom-in duration-200">
        <h3 className="text-xl font-bold text-slate-900 mb-2">Enable AI Enhancement</h3>
        <p className="text-slate-500 mb-6 text-sm">
          To generate speaker notes, please enter your Google Gemini API Key.
          <br/>
          <a 
            href="https://aistudio.google.com/app/apikey" 
            target="_blank" 
            rel="noopener noreferrer"
            className="text-blue-600 hover:text-blue-700 hover:underline"
          >
            Get an API key from Google AI Studio
          </a>
        </p>
        
        <input
          type="password"
          value={key}
          onChange={(e) => setKey(e.target.value)}
          placeholder="Enter your API Key"
          className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none mb-6 font-mono text-sm"
          autoFocus
        />

        <div className="flex justify-end space-x-3">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-slate-600 hover:bg-slate-100 rounded-lg font-medium transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={() => onSave(key)}
            disabled={!key.trim()}
            className="px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            Enable AI
          </button>
        </div>
      </div>
    </div>
  );
};