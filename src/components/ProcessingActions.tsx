import React from 'react';
import { FileUp, Download, Code, RefreshCw, Loader2 } from 'lucide-react';
import { VBAModule } from '../utils/vbaCodeExtractor/types';

interface ProcessingActionsProps {
  file: File | null;
  isProcessing: boolean;
  processedFile: Blob | null;
  extractedModules: VBAModule[];
  progress: number;
  onRemovePassword: () => void;
  onExtractCode: () => void;
  onDownloadFile: () => void;
  onDownloadVBACode: () => void;
  onReset: () => void;
}

export function ProcessingActions({
  file,
  isProcessing,
  processedFile,
  extractedModules,
  progress,
  onRemovePassword,
  onExtractCode,
  onDownloadFile,
  onDownloadVBACode,
  onReset
}: ProcessingActionsProps) {
  if (!file) {
    return null;
  }

  return (
    <div className="mt-6">
      <h3 className="text-lg font-medium text-gray-900 mb-4">Actions</h3>
      
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <button
          type="button"
          onClick={onRemovePassword}
          disabled={isProcessing}
          className={`
            flex items-center justify-center px-4 py-2 border border-transparent 
            text-sm font-medium rounded-md shadow-sm text-white 
            ${isProcessing ? 'bg-gray-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700'}
            focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500
          `}
        >
          {isProcessing ? (
            <>
              <Loader2 className="animate-spin -ml-1 mr-2 h-4 w-4" />
              Processing... {progress > 0 ? `${progress}%` : ''}
            </>
          ) : (
            <>
              <FileUp className="-ml-1 mr-2 h-4 w-4" />
              Remove VBA Password
            </>
          )}
        </button>
        
        <button
          type="button"
          onClick={onExtractCode}
          disabled={isProcessing}
          className={`
            flex items-center justify-center px-4 py-2 border border-transparent 
            text-sm font-medium rounded-md shadow-sm text-white 
            ${isProcessing ? 'bg-gray-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700'}
            focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500
          `}
        >
          {isProcessing ? (
            <>
              <Loader2 className="animate-spin -ml-1 mr-2 h-4 w-4" />
              Extracting... {progress > 0 ? `${progress}%` : ''}
            </>
          ) : (
            <>
              <Code className="-ml-1 mr-2 h-4 w-4" />
              Extract VBA Code
            </>
          )}
        </button>
        
        {processedFile && (
          <button
            type="button"
            onClick={onDownloadFile}
            className="flex items-center justify-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
          >
            <Download className="-ml-1 mr-2 h-4 w-4" />
            Download Unprotected File
          </button>
        )}
        
        {extractedModules.length > 0 && (
          <button
            type="button"
            onClick={onDownloadVBACode}
            className="flex items-center justify-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
          >
            <Code className="-ml-1 mr-2 h-4 w-4" />
            Download VBA Code
          </button>
        )}
        
        <button
          type="button"
          onClick={onReset}
          className="flex items-center justify-center px-4 py-2 border border-gray-300 text-sm font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
        >
          <RefreshCw className="-ml-1 mr-2 h-4 w-4" />
          Reset
        </button>
      </div>
      
      {isProcessing && (
        <div className="mt-4">
          <div className="w-full bg-gray-200 rounded-full h-2.5">
            <div 
              className="bg-indigo-600 h-2.5 rounded-full" 
              style={{ width: `${progress}%` }}
            ></div>
          </div>
        </div>
      )}
    </div>
  );
} 