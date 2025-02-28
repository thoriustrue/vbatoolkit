import React, { useCallback, useRef, useState } from 'react';
import { Upload } from 'lucide-react';

interface FileUploaderProps {
  onFileSelect: (file: File) => void;
  acceptedExtensions: string[];
  maxSizeInMB: number;
  addLog: (message: string, type: 'info' | 'error' | 'success') => void;
}

export function FileUploader({ 
  onFileSelect, 
  acceptedExtensions, 
  maxSizeInMB, 
  addLog 
}: FileUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validateFile = (file: File): boolean => {
    const fileExtension = '.' + file.name.split('.').pop()?.toLowerCase();
    
    if (!acceptedExtensions.includes(fileExtension)) {
      addLog(`Invalid file type: ${fileExtension}. Please upload ${acceptedExtensions.join(', ')} files.`, 'error');
      return false;
    }
    
    if (file.size > maxSizeInMB * 1024 * 1024) {
      addLog(`File is too large. Maximum size is ${maxSizeInMB}MB.`, 'error');
      return false;
    }
    
    return true;
  };

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  const handleFileDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (validateFile(droppedFile)) {
        onFileSelect(droppedFile);
        addLog(`File selected: ${droppedFile.name} (${(droppedFile.size / 1024).toFixed(2)} KB)`, 'info');
      }
    }
  }, [addLog, onFileSelect, validateFile]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFile = e.target.files[0];
      if (validateFile(selectedFile)) {
        onFileSelect(selectedFile);
        addLog(`File selected: ${selectedFile.name} (${(selectedFile.size / 1024).toFixed(2)} KB)`, 'info');
      }
    }
  }, [addLog, onFileSelect, validateFile]);

  const handleButtonClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  return (
    <div
      className={`border-2 border-dashed rounded-lg p-6 text-center ${
        isDragging ? 'border-indigo-500 bg-indigo-50' : 'border-gray-300'
      }`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleFileDrop}
    >
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileSelect}
        className="hidden"
        accept={acceptedExtensions.join(',')}
      />
      <Upload className="mx-auto h-12 w-12 text-gray-400" />
      <h3 className="mt-2 text-sm font-medium text-gray-900">
        Drag and drop your Excel file here
      </h3>
      <p className="mt-1 text-xs text-gray-500">
        Or
      </p>
      <button
        type="button"
        onClick={handleButtonClick}
        className="mt-2 inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
      >
        Select File
      </button>
      <p className="mt-2 text-xs text-gray-500">
        Supported formats: {acceptedExtensions.join(', ')} (Max: {maxSizeInMB}MB)
      </p>
    </div>
  );
} 