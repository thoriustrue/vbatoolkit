import React, { useState, useCallback } from 'react';
import { Upload } from 'lucide-react';
import { removeVBAPassword } from './utils/vbaPasswordRemover';
import { extractVBACode, VBAModule, createVBACodeFile } from './utils/vbaCodeExtractor/index';
import { injectVBACode } from './utils/vbaCodeInjector';
import { ErrorBoundary, useErrorLogger } from './components/ErrorLogger';
import { ErrorLog } from './components/ErrorLog';
import { FileUploader } from './components/FileUploader';
import { LogViewer } from './components/LogViewer';
import { ProcessingActions } from './components/ProcessingActions';
import { Changelog, ChangelogEntry as ChangelogEntryComponent } from './components/Changelog';
import { CHANGELOG_DATA } from './components/Changelog';
import { LogEntry, LogType, ChangelogChange } from './types';

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedFile, setProcessedFile] = useState<Blob | null>(null);
  const [extractedModules, setExtractedModules] = useState<VBAModule[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [progress, setProgress] = useState(0);
  const [activeTab, setActiveTab] = useState('main');
  const { logError } = useErrorLogger();

  const addLog = useCallback((message: string, type: LogType = 'info') => {
    setLogs(prev => [...prev, { message, type }]);
  }, []);

  const resetProcess = useCallback(() => {
    setFile(null);
    setProcessedFile(null);
    setExtractedModules([]);
    setLogs([]);
    setProgress(0);
  }, []);

  const handleFileSelect = useCallback((selectedFile: File) => {
    setFile(selectedFile);
    setProcessedFile(null);
    setExtractedModules([]);
    setProgress(0);
    setLogs([]);
  }, []);

  const removePassword = useCallback(async () => {
    if (!file) return;
    
    setIsProcessing(true);
    setLogs([]);
    setProgress(0);
    addLog('Starting VBA password removal process...', 'info');
    
    try {
      const result = await removeVBAPassword(file, (message: string, type: LogType) => {
        addLog(message, type);
      }, (progressValue: number) => {
        setProgress(progressValue * 100);
      });
      
      if (result) {
        setProcessedFile(result);
        addLog('VBA password removal completed successfully!', 'success');
      } else {
        addLog('Failed to remove VBA password. See errors above.', 'error');
      }
    } catch (error) {
      logError(error instanceof Error ? error : new Error(String(error)));
      addLog(`Error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    } finally {
      setIsProcessing(false);
      setProgress(100);
    }
  }, [file, addLog, logError]);

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
                  <h2 className="text-xl font-semibold text-gray-900 mb-2">VBA Password Remover & Code Extractor</h2>
                  <p className="text-gray-600">
                    Upload an Excel file with VBA macros to remove password protection or extract the VBA code.
                  </p>
                </div>
                
                {!file ? (
                  <FileUploader 
                    onFileSelect={handleFileSelect}
                    acceptedExtensions={['.xlsm', '.xls', '.xlsb']}
                    maxSizeInMB={50}
                    addLog={addLog}
                  />
                ) : (
                  <>
                    <div className="bg-blue-50 border border-blue-200 rounded-md p-4 mb-4">
                      <p className="text-sm text-blue-800">
                        <strong>Selected file:</strong> {file.name} ({(file.size / 1024).toFixed(2)} KB)
                      </p>
                    </div>
                    
                    <ProcessingActions
                      file={file}
                      isProcessing={isProcessing}
                      processedFile={processedFile}
                      extractedModules={extractedModules}
                      progress={progress}
                      onRemovePassword={removePassword}
                      onExtractCode={extractCode}
                      onDownloadFile={downloadFile}
                      onDownloadVBACode={downloadVBACode}
                      onReset={resetProcess}
                    />
                    
                    <LogViewer logs={logs} onClearLogs={clearLogs} />
                  </>
                )}
                
                <Changelog>
                  {CHANGELOG_DATA.map((entry) => (
                    <ChangelogEntryComponent key={entry.version} entry={entry} />
                  ))}
                </Changelog>
              </div>
            </div>
          </div>
        </main>
        
        <footer className="bg-white mt-8 py-4 border-t">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <p className="text-center text-sm text-gray-500">
              Excel VBA Tools - For educational purposes only. Use responsibly.
            </p>
          </div>
        </footer>
      </div>
    </ErrorBoundary>
  );
}

export default App;