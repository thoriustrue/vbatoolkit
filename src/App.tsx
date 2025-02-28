import React, { useState, useRef, useCallback } from 'react';
import { Upload, FileUp, Download, AlertCircle, CheckCircle, Info, Loader2, RefreshCw, Code, Shield, ChevronDown, ChevronUp } from 'lucide-react';
import { removeVBAPassword } from './utils/vbaPasswordRemover';
import { extractVBACode, VBAModule, createVBACodeFile } from './utils/vbaCodeExtractor';
import { injectVBACode } from './utils/vbaCodeInjector';
import { ErrorBoundary, ErrorLogPanel, useErrorLogger } from './components/ErrorLogger';
import { ErrorLog } from './components/ErrorLog';

// Define changelog data directly in App.tsx to avoid import issues
interface ChangelogEntry {
  version: string;
  date: string;
  changes: {
    type: 'added' | 'fixed' | 'changed' | 'removed';
    description: string;
  }[];
}

const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '0.1.1',
    date: '2024-06-20',
    changes: [
      { type: 'added', description: 'Added VBA password removal functionality' },
      { type: 'added', description: 'Implemented macro auto-enable features' },
      { type: 'added', description: 'Added error logging system' }
    ]
  },
  {
    version: '0.1.0',
    date: '2024-06-15',
    changes: [
      { type: 'added', description: 'Initial release of VBA Toolkit' },
      { type: 'added', description: 'Basic file handling capabilities' },
      { type: 'added', description: 'User interface for file operations' }
    ]
  }
];

// Simple Changelog component defined directly in App.tsx
function Changelog({ children }: { children?: React.ReactNode }) {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <details className="mt-4 text-sm text-gray-600 dark:text-gray-300">
      <summary 
        className="flex items-center cursor-pointer list-none"
        onClick={(e) => {
          e.preventDefault();
          setIsOpen(!isOpen);
        }}
      >
        {isOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
        <span className="ml-2">Version History</span>
      </summary>
      <div className="ml-6 mt-2 space-y-2">
        {children}
      </div>
    </details>
  );
}

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedFile, setProcessedFile] = useState<Blob | null>(null);
  const [extractedModules, setExtractedModules] = useState<VBAModule[]>([]);
  const [logs, setLogs] = useState<Array<{ message: string; type: 'info' | 'error' | 'success' }>>([]);
  const [progress, setProgress] = useState(0);
  const [activeTab, setActiveTab] = useState('main');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { logError } = useErrorLogger();

  const addLog = useCallback((message: string, type: 'info' | 'error' | 'success' = 'info') => {
    setLogs(prev => [...prev, { message, type }]);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  const validateFile = (file: File): boolean => {
    const validExtensions = ['.xlsm', '.xls', '.xlsb'];
    const fileExtension = '.' + file.name.split('.').pop()?.toLowerCase();
    
    if (!validExtensions.includes(fileExtension)) {
      addLog(`Invalid file type: ${fileExtension}. Please upload .xlsm, .xls, or .xlsb files.`, 'error');
      return false;
    }
    
    if (file.size > 50 * 1024 * 1024) { // 50MB limit
      addLog('File is too large. Maximum size is 50MB.', 'error');
      return false;
    }
    
    return true;
  };

  const handleFileDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (validateFile(droppedFile)) {
        setFile(droppedFile);
        addLog(`File selected: ${droppedFile.name} (${(droppedFile.size / 1024).toFixed(2)} KB)`, 'info');
        setProcessedFile(null);
        setExtractedModules([]);
        setProgress(0);
        setLogs([]);
      }
    }
  }, [addLog]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFile = e.target.files[0];
      if (validateFile(selectedFile)) {
        setFile(selectedFile);
        addLog(`File selected: ${selectedFile.name} (${(selectedFile.size / 1024).toFixed(2)} KB)`, 'info');
        setProcessedFile(null);
        setExtractedModules([]);
        setProgress(0);
        setLogs([]);
      }
    }
  }, [addLog]);

  const handleButtonClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const resetProcess = useCallback(() => {
    setProcessedFile(null);
    setExtractedModules([]);
    setProgress(0);
    setLogs([]);
    addLog(`Ready to process file: ${file?.name}`, 'info');
  }, [file, addLog]);

  const handleFileUpload = async (file: File) => {
    if (!validateFile(file)) return;

    setIsProcessing(true);
    addLog('Processing file...', 'info');

    try {
      const fileData = await file.arrayBuffer();
      if (activeTab === 'alternate') {
        await injectVBACode(fileData, addLog);
      } else {
        // Existing processing logic
      }
    } catch (error) {
      addLog(`Error processing file: ${error.message}`, 'error');
      logError(error, 'fileProcessing'); // Log to technical error panel
    } finally {
      setIsProcessing(false);
    }
  };

  const processFile = useCallback(async () => {
    if (!file) return;
    
    setIsProcessing(true);
    setLogs([]);
    setProgress(0);
    addLog('Starting VBA password removal process...', 'info');
    addLog('Auto-enabling macros and external links...', 'info');
    
    try {
      const result = await removeVBAPassword(file, (message, type) => {
        addLog(message, type);
      }, (progressValue) => {
        setProgress(progressValue);
      });
      
      if (result) {
        setProcessedFile(result);
        addLog('VBA password removal completed successfully!', 'success');
        addLog(`Original file size: ${(file.size / 1024).toFixed(2)} KB, Processed file size: ${(result.size / 1024).toFixed(2)} KB`, 'info');
      } else {
        addLog('Failed to process the file. See errors above.', 'error');
      }
    } catch (error) {
      addLog(`Error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    } finally {
      setIsProcessing(false);
      setProgress(100);
    }
  }, [file, addLog]);

  const extractCode = useCallback(async () => {
    if (!file) return;
    
    setIsProcessing(true);
    setLogs([]);
    setProgress(0);
    addLog('Starting VBA code extraction process...', 'info');
    
    try {
      const result = await extractVBACode(file, (message, type) => {
        addLog(message, type);
      }, (progressValue) => {
        setProgress(progressValue);
      });
      
      if (result.success && result.modules.length > 0) {
        setExtractedModules(result.modules);
        addLog(`Successfully extracted ${result.modules.length} VBA module(s)!`, 'success');
        
        // Log each module
        result.modules.forEach(module => {
          addLog(`Module: ${module.name} (${module.type})`, 'info');
        });
      } else {
        addLog('Failed to extract VBA code. See errors above.', 'error');
      }
    } catch (error) {
      addLog(`Error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    } finally {
      setIsProcessing(false);
      setProgress(100);
    }
  }, [file, addLog]);

  const downloadFile = useCallback(() => {
    if (!processedFile || !file) return;
    
    const fileName = file.name;
    const fileExtension = '.' + fileName.split('.').pop();
    const newFileName = fileName.replace(fileExtension, `_unprotected${fileExtension}`);
    
    // Create Blob and download using native API
    const blob = new Blob([processedFile], { 
      type: 'application/vnd.ms-excel.sheet.macroEnabled.12' 
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = newFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    addLog(`File downloaded as: ${newFileName}`, 'success');
  }, [processedFile, file, addLog]);

  const downloadVBACode = useCallback(() => {
    if (extractedModules.length === 0 || !file) return;
    
    const codeFile = createVBACodeFile(extractedModules, file.name);
    const fileName = file.name.split('.')[0] + '_vba_code.txt';
    
    const url = URL.createObjectURL(codeFile);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    addLog(`VBA code downloaded as: ${fileName}`, 'success');
  }, [extractedModules, file, addLog]);

  const getLogIcon = (type: 'info' | 'error' | 'success') => {
    switch (type) {
      case 'info': return <Info className="w-4 h-4 text-blue-500" />;
      case 'error': return <AlertCircle className="w-4 h-4 text-red-500" />;
      case 'success': return <CheckCircle className="w-4 h-4 text-green-500" />;
    }
  };

  const clearLogs = useCallback(() => {
    setLogs([]);
  }, []);

  return (
    <ErrorBoundary>
      <div className="min-h-screen bg-gray-50 p-4">
        <header className="bg-white shadow-sm">
          <div className="max-w-7xl mx-auto py-4 px-4 sm:px-6 lg:px-8 flex items-center">
            <Upload className="h-8 w-8 text-indigo-600 mr-3" />
            <h1 className="text-2xl font-bold text-gray-900">Excel VBA Tools</h1>
          </div>
        </header>
        
        <main className="flex-grow">
          <div className="max-w-7xl mx-auto py-6 sm:px-6 lg:px-8">
            <div className="px-4 py-6 sm:px-0">
              <div className="bg-white rounded-lg shadow p-6">
                <div className="mb-6">
                  <h2 className="text-lg font-medium text-gray-900 mb-2">Upload Excel File</h2>
                  <p className="text-sm text-gray-500 mb-4">
                    Upload an Excel file (.xlsm, .xls, .xlsb) with VBA code to remove password protection or extract the VBA code.
                    All processing happens in your browser - no files are sent to any server.
                  </p>
                  
                  <div 
                    className={`border-2 border-dashed rounded-lg p-8 text-center ${
                      isDragging ? 'border-indigo-500 bg-indigo-50' : 'border-gray-300'
                    } ${file ? 'bg-green-50' : ''}`}
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                    onDrop={handleFileDrop}
                  >
                    <input
                      type="file"
                      ref={fileInputRef}
                      className="hidden"
                      accept=".xlsm,.xls,.xlsb"
                      onChange={handleFileSelect}
                    />
                    
                    <div className="space-y-2">
                      <div className="flex justify-center">
                        <FileUp className="h-12 w-12 text-gray-400" />
                      </div>
                      <div className="text-sm text-gray-600">
                        {file ? (
                          <p className="font-medium text-green-600">{file.name} selected ({(file.size / 1024).toFixed(2)} KB)</p>
                        ) : (
                          <>
                            <p className="font-medium">Drag and drop your Excel file here</p>
                            <p>or</p>
                          </>
                        )}
                      </div>
                      {!file && (
                        <button
                          type="button"
                          onClick={handleButtonClick}
                          className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                        >
                          Select File
                        </button>
                      )}
                    </div>
                  </div>
                </div>
                
                <div className="flex border-b border-gray-200 dark:border-gray-700">
                  <div className="px-4 py-2 border-b-2 border-blue-500 text-blue-600 dark:text-blue-400">
                    VBA Password Remover
                  </div>
                </div>
                
                <div className="p-4">
                  {/* Main method content */}
                  <div className="space-y-4">
                    <div className="flex items-center space-x-2">
                      <Upload className="text-blue-500" size={20} />
                      <h2 className="text-lg font-medium">Upload Excel File</h2>
                    </div>
                    
                    {/* File upload section */}
                    <div 
                      className={`border-2 border-dashed rounded-lg p-8 text-center ${
                        isDragging ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20' : 'border-gray-300 dark:border-gray-700'
                      }`}
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleFileDrop}
                    >
                      {/* Keep all the original content from the main tab here */}
                      {file ? (
                        <div className="space-y-2">
                          <div className="flex items-center justify-center space-x-2">
                            <FileUp className="text-green-500" size={24} />
                            <span className="font-medium">{file.name}</span>
                            <span className="text-sm text-gray-500">
                              ({(file.size / 1024).toFixed(2)} KB)
                            </span>
                          </div>
                          <button
                            className="px-3 py-1 bg-red-100 text-red-700 rounded hover:bg-red-200 dark:bg-red-900/30 dark:text-red-400 dark:hover:bg-red-900/50 transition-colors"
                            onClick={() => {
                              setFile(null);
                              setIsProcessing(false);
                              setProcessedFile(null);
                              setProgress(0);
                            }}
                          >
                            Remove
                          </button>
                        </div>
                      ) : (
                        <div className="space-y-4">
                          <p className="text-gray-500 dark:text-gray-400">
                            Drag and drop your Excel file here, or click to select
                          </p>
                          <button
                            className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 dark:bg-blue-600 dark:hover:bg-blue-700 transition-colors"
                            onClick={() => fileInputRef.current?.click()}
                          >
                            Select File
                          </button>
                          <input
                            ref={fileInputRef}
                            type="file"
                            accept=".xlsm,.xlsb,.xls,.xlam"
                            className="hidden"
                            onChange={handleFileSelect}
                          />
                        </div>
                      )}
                    </div>
                    
                    {/* Keep all the processing section from the main tab */}
                    {file && (
                      <div className="space-y-4">
                        <div className="flex items-center space-x-2">
                          <Shield className="text-blue-500" size={20} />
                          <h2 className="text-lg font-medium">Remove VBA Password</h2>
                        </div>
                        
                        <div className="p-4 border rounded-lg dark:border-gray-700">
                          <div className="space-y-4">
                            <p className="text-gray-600 dark:text-gray-300">
                              Click the button below to remove the VBA password protection from your Excel file.
                            </p>
                            
                            {!isProcessing && !processedFile && (
                              <button
                                className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 dark:bg-blue-600 dark:hover:bg-blue-700 transition-colors"
                                onClick={processFile}
                              >
                                Remove Password
                              </button>
                            )}
                            
                            {isProcessing && (
                              <div className="space-y-2">
                                <div className="flex items-center space-x-2">
                                  <Loader2 className="animate-spin text-blue-500" size={20} />
                                  <span className="text-blue-600 dark:text-blue-400">Processing...</span>
                                </div>
                                <div className="w-full bg-gray-200 rounded-full h-2.5 dark:bg-gray-700">
                                  <div 
                                    className="bg-blue-500 h-2.5 rounded-full transition-all duration-300" 
                                    style={{ width: `${progress * 100}%` }}
                                  ></div>
                                </div>
                              </div>
                            )}
                            
                            {processedFile && (
                              <div className="space-y-4">
                                <div className="flex items-center space-x-2">
                                  <CheckCircle className="text-green-500" size={20} />
                                  <span className="text-green-600 dark:text-green-400">Password successfully removed!</span>
                                </div>
                                
                                <div className="flex space-x-2">
                                  <button
                                    className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 dark:bg-green-600 dark:hover:bg-green-700 transition-colors flex items-center space-x-2"
                                    onClick={downloadFile}
                                  >
                                    <Download size={16} />
                                    <span>Download Unprotected File</span>
                                  </button>
                                  
                                  <button
                                    className="px-4 py-2 bg-blue-100 text-blue-700 rounded hover:bg-blue-200 dark:bg-blue-900/30 dark:text-blue-400 dark:hover:bg-blue-900/50 transition-colors flex items-center space-x-2"
                                    onClick={() => {
                                      setIsProcessing(false);
                                      setProcessedFile(null);
                                      setProgress(0);
                                    }}
                                  >
                                    <RefreshCw size={16} />
                                    <span>Process Again</span>
                                  </button>
                                </div>
                                
                                <div className="text-sm text-gray-500 dark:text-gray-400">
                                  Original file size: {(file.size / 1024).toFixed(2)} KB, 
                                  Processed file size: {(processedFile.size / 1024).toFixed(2)} KB
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    )}
                    
                    {/* VBA Code Extraction Section */}
                    {file && (
                      <div className="space-y-4">
                        <div className="flex items-center space-x-2">
                          <Code className="text-blue-500" size={20} />
                          <h2 className="text-lg font-medium">Extract VBA Code</h2>
                        </div>
                        
                        <div className="p-4 border rounded-lg dark:border-gray-700">
                          <div className="space-y-4">
                            <p className="text-gray-600 dark:text-gray-300">
                              Extract all VBA code from the Excel file for review or backup.
                            </p>
                            
                            {!isProcessing && !extractedModules && (
                              <button
                                className="px-4 py-2 bg-purple-500 text-white rounded hover:bg-purple-600 dark:bg-purple-600 dark:hover:bg-purple-700 transition-colors"
                                onClick={extractCode}
                              >
                                Extract VBA Code
                              </button>
                            )}
                            
                            {isProcessing && (
                              <div className="flex items-center space-x-2">
                                <Loader2 className="animate-spin text-purple-500" size={20} />
                                <span className="text-purple-600 dark:text-purple-400">Extracting code...</span>
                              </div>
                            )}
                            
                            {extractedModules && (
                              <div className="space-y-4">
                                <div className="flex items-center space-x-2">
                                  <CheckCircle className="text-green-500" size={20} />
                                  <span className="text-green-600 dark:text-green-400">
                                    Successfully extracted {extractedModules.length} VBA modules!
                                  </span>
                                </div>
                                
                                <button
                                  className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 dark:bg-green-600 dark:hover:bg-green-700 transition-colors flex items-center space-x-2"
                                  onClick={downloadVBACode}
                                >
                                  <Download size={16} />
                                  <span>Download VBA Code</span>
                                </button>
                                
                                <div className="space-y-2">
                                  <h3 className="font-medium">Extracted Modules:</h3>
                                  <ul className="list-disc pl-5 space-y-1">
                                    {extractedModules.map((module, index) => (
                                      <li key={index} className="text-sm">
                                        {module.name} ({module.type})
                                      </li>
                                    ))}
                                  </ul>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </main>
        
        {/* Footer with changelog */}
        <footer className="mt-auto p-4 border-t border-gray-200 dark:border-gray-700">
          <Changelog>
            {CHANGELOG_DATA.map((entry, index) => (
              <div key={index} className="space-y-2 mb-4">
                <h3 className="font-medium">v{entry.version} - {entry.date}</h3>
                <ul className="list-disc pl-4 space-y-1">
                  {entry.changes.map((change, i) => (
                    <li key={i} className={
                      change.type === 'added' ? 'text-green-600 dark:text-green-400' :
                      change.type === 'fixed' ? 'text-amber-600 dark:text-amber-400' :
                      'text-blue-600 dark:text-blue-400'
                    }>
                      {change.description}
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </Changelog>
        </footer>
        
        <ErrorLog logs={logs} onClear={clearLogs} />
      </div>
    </ErrorBoundary>
  );
}

export default App;