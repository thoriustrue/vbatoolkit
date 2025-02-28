import React, { useState, useRef, useCallback } from 'react';
import { Upload, FileUp, Download, AlertCircle, CheckCircle, Info, Loader2, RefreshCw, Code, Shield } from 'lucide-react';
import { removeVBAPassword } from './utils/vbaPasswordRemover';
import { extractVBACode, VBAModule, createVBACodeFile } from './utils/vbaCodeExtractor';
import { injectVBACode } from './utils/vbaCodeInjector';
import { ErrorBoundary, ErrorLogPanel, useErrorLogger } from './components/ErrorLogger';
import { Changelog } from './components/Changelog';
import { ErrorLog } from './components/ErrorLog';
import { CHANGELOG_DATA } from './data/changelog';

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedFile, setProcessedFile] = useState<Blob | null>(null);
  const [extractedModules, setExtractedModules] = useState<VBAModule[]>([]);
  const [logs, setLogs] = useState<Array<{ message: string; type: 'info' | 'error' | 'success' }>>([]);
  const [progress, setProgress] = useState(0);
  const [activeTab, setActiveTab] = useState<'remove' | 'extract' | 'alternative'>('remove');
  const [method, setMethod] = useState<'remove' | 'extract' | 'alternative'>('remove');
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
      if (method === 'alternative') {
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
                
                {file && (
                  <div className="mt-6 space-y-4">
                    {/* Tabs */}
                    <div className="border-b border-gray-200">
                      <nav className="-mb-px flex" aria-label="Tabs">
                        <button
                          onClick={() => setActiveTab('remove')}
                          className={`${
                            activeTab === 'remove'
                              ? 'border-indigo-500 text-indigo-600'
                              : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                          } w-1/3 py-4 px-1 text-center border-b-2 font-medium text-sm`}
                        >
                          Remove VBA Password
                        </button>
                        <button
                          onClick={() => setActiveTab('extract')}
                          className={`${
                            activeTab === 'extract'
                              ? 'border-indigo-500 text-indigo-600'
                              : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                          } w-1/3 py-4 px-1 text-center border-b-2 font-medium text-sm`}
                        >
                          Extract VBA Code
                        </button>
                        <button
                          onClick={() => setActiveTab('alternative')}
                          className={`${
                            activeTab === 'alternative'
                              ? 'border-indigo-500 text-indigo-600'
                              : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                          } w-1/3 py-4 px-1 text-center border-b-2 font-medium text-sm`}
                        >
                          Alternative Method
                        </button>
                      </nav>
                    </div>
                    
                    <div className="flex justify-between items-center">
                      <h3 className="text-md font-medium text-gray-900">
                        {activeTab === 'remove' ? 'Remove VBA Password' : activeTab === 'extract' ? 'Extract VBA Code' : 'Alternative Method'}
                      </h3>
                      <div className="flex space-x-2">
                        {!isProcessing && activeTab === 'remove' && !processedFile && (
                          <button
                            type="button"
                            onClick={processFile}
                            className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo -700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                          >
                            <Shield className="h-4 w-4 mr-2 " />
                            Remove VBA Password
                          </button>
                        )}
                        
                        {!isProcessing && activeTab === 'extract' && extractedModules.length === 0 && (
                          <button
                            type="button"
                            onClick={extractCode}
                            className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                          >
                            <Code className="h-4 w-4 mr-2" />
                            Extract VBA Code
                          </button>
                        )}
                        
                        {activeTab === 'remove' && processedFile && (
                          <>
                            <button
                              type="button"
                              onClick={resetProcess}
                              className="inline-flex items-center px-4 py-2 border border-gray-300 text-sm font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                            >
                              <RefreshCw className="h-4 w-4 mr-2" />
                              Try Again
                            </button>
                            <button
                              type="button"
                              onClick={downloadFile}
                              className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
                            >
                              <Download className="h-4 w-4 mr-2" />
                              Download Unprotected File
                            </button>
                          </>
                        )}
                        
                        {activeTab === 'extract' && extractedModules.length > 0 && (
                          <>
                            <button
                              type="button"
                              onClick={resetProcess}
                              className="inline-flex items-center px-4 py-2 border border-gray-300 text-sm font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                            >
                              <RefreshCw className="h-4 w-4 mr-2" />
                              Try Again
                            </button>
                            <button
                              type="button"
                              onClick={downloadVBACode}
                              className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
                            >
                              <Download className="h-4 w-4 mr-2" />
                              Download VBA Code
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                    
                    {/* Progress bar */}
                    {(isProcessing || progress > 0) && (
                      <div className="mt-2">
                        <div className="relative pt-1">
                          <div className="flex mb-2 items-center justify-between">
                            <div>
                              <span className="text-xs font-semibold inline-block py-1 px-2 uppercase rounded-full text-indigo-600 bg-indigo-200">
                                Progress
                              </span>
                            </div>
                            <div className="text-right">
                              <span className="text-xs font-semibold inline-block text-indigo-600">
                                {progress}%
                              </span>
                            </div>
                          </div>
                          <div className="overflow-hidden h-2 mb-4 text-xs flex rounded bg-indigo-200">
                            <div 
                              style={{ width: `${progress}%` }} 
                              className="shadow-none flex flex-col text-center whitespace-nowrap text-white justify-center bg-indigo-500 transition-all duration-300"
                            ></div>
                          </div>
                        </div>
                      </div>
                    )}
                    
                    {/* Extracted modules list */}
                    {activeTab === 'extract' && extractedModules.length > 0 && (
                      <div className="mt-4">
                        <h3 className="text-md font-medium text-gray-900 mb-2">Extracted VBA Modules</h3>
                        <div className="bg-gray-50 rounded-md p-3 border border-gray-200 max-h-64 overflow-y-auto">
                          <ul className="divide-y divide-gray-200">
                            {extractedModules.map((module, index) => (
                              <li key={index} className="py-3">
                                <div className="flex items-center justify-between">
                                  <div>
                                    <p className="text-sm font-medium text-gray-900">{module.name}</p>
                                    <p className="text-sm text-gray-500">{module.type}</p>
                                  </div>
                                  <div className="text-sm text-gray-500">
                                    {(module.code.length / 1024).toFixed(2)} KB
                                  </div>
                                </div>
                              </li>
                            ))}
                          </ul>
                        </div>
                      </div>
                    )}
                    
                    {/* Process log */}
                    <div className="mt-4">
                      <h3 className="text-md font-medium text-gray-900 mb-2">Process Log</h3>
                      <div className="bg-gray-50 rounded-md p-3 h-64 overflow-y-auto border border-gray-200">
                        {logs.length === 0 ? (
                          <p className="text-gray-500 text-sm italic">Process logs will appear here...</p>
                        ) : (
                          <div className="space-y-1">
                            {logs.map((log, index) => (
                              <div key={index} className="flex items-start text-sm">
                                <span className="mr-2 mt-0.5">{getLogIcon(log.type)}</span>
                                <span className={`${
                                  log.type === 'error' ? 'text-red-600' : 
                                  log.type === 'success' ? 'text-green-600' : 'text-gray-700'
                                }`}>
                                  {log.message}
                                </span>
                              </div>
                            ))}
                            {isProcessing && (
                              <div className="flex items-center text-sm text-indigo-600">
                                <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                                <span>Processing...</span>
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                )}
                
                {/* Feature descriptions */}
                <div className="mt-8 grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="p-4 bg-blue-50 rounded-md border border-blue-100">
                    <h3 className="text-md font-medium text-blue-800 mb-2 flex items-center">
                      <Shield className="w-5 h-5 mr-2" />
                      Remove VBA Password
                    </h3>
                    <p className="text-sm text-blue-700 mb-2">
                      This tool removes password protection from VBA projects in Excel files, allowing you to access and edit the VBA code without knowing the original password.
                    </p>
                    <ul className="text-sm text-blue-700 list-disc pl-5 space-y-1">
                      <li>Works with Excel 2007-2022 files (.xlsm, .xls, .xlsb)</li>
                      <li>Removes project-level password protection</li>
                      <li>Removes sheet and workbook protection</li>
                      <li>Auto-enables macros and external links</li>
                      <li>100% client-side processing (no data is sent to any server)</li>
                    </ul>
                  </div>
                  
                  <div className="p-4 bg-indigo-50 rounded-md border border-indigo-100">
                    <h3 className="text-md font-medium text-indigo-800 mb-2 flex items-center">
                      <Code className="w-5 h-5 mr-2" />
                      Extract VBA Code
                    </h3>
                    <p className="text-sm text-indigo-700 mb-2">
                      This tool extracts all VBA code modules from Excel files, allowing you to view, backup, or reuse the code without opening Excel.
                    </p>
                    <ul className="text-sm text-indigo-700 list-disc pl-5 space-y-1">
                      <li>Extracts all types of VBA modules (standard, class, form, document)</li>
                      <li>Preserves module names and types</li>
                      <li>Exports code to a plain text file for easy viewing or backup</li>
                      <li>Works with both protected and unprotected VBA projects</li>
                    </ul>
                  </div>
                </div>
                
                {/* Disclaimer */}
                <div className="mt-8 p-4 bg-yellow-50 rounded-md border border-yellow-100">
                  <h3 className="text-md font-medium text-yellow-800 mb-2 flex items-center">
                    <AlertCircle className="w-5 h-5 mr-2" />
                    Ethical Usage Disclaimer
                  </h3>
                  <p className="text-sm text-yellow-700">
                    These tools should only be used on Excel files that you own or have explicit permission to modify.
                    They are intended for legitimate purposes, such as accessing your own VBA projects
                    when you've forgotten the password or backing up your code. Unauthorized access to protected files may violate applicable laws.
                  </p>
                </div>
                
                {/* How it works */}
                <div className="mt-8">
                  <h3 className="text-md font-medium text-gray-900 mb-2">How It Works</h3>
                  <div className="text-sm text-gray-600 space-y-2">
                    <p>
                      <strong>Password Removal:</strong> This tool works by manipulating the binary structure of the Excel VBA project to remove password protection.
                      It locates the password hash in the VBA project structure and removes it, allowing you to access the VBA code without needing the original password.
                      It also removes sheet and workbook protection and configures security settings to auto-enable macros and external links.
                    </p>
                    <p>
                      <strong>Code Extraction:</strong> This tool analyzes the VBA project structure in the Excel file and extracts all code modules.
                      It identifies different module types (standard modules, class modules, forms, and document modules) and exports them to a text file.
                    </p>
                    <p>
                      The entire process happens in your browser - no data is sent to any server.
                    </p>
                    <p>
                      Supported Excel versions: Excel 2007-2022 (.xlsm, .xls, .xlsb formats)
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </main>
        
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