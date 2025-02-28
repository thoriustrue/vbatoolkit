import { useState, useEffect } from 'react';
import { AlertCircle, ChevronDown, ChevronUp, Copy, X } from 'lucide-react';

interface ErrorLogEntry {
  id: string;
  timestamp: Date;
  message: string;
  stack?: string;
  componentStack?: string;
  source?: string;
}

export function useErrorLogger() {
  const [errors, setErrors] = useState<ErrorLogEntry[]>([]);
  
  const logError = (error: Error | string, source = 'app') => {
    const newError: ErrorLogEntry = {
      id: crypto.randomUUID(),
      timestamp: new Date(),
      message: error instanceof Error ? error.message : error,
      stack: error instanceof Error ? error.stack : undefined,
      source
    };
    
    setErrors(prev => [newError, ...prev]);
    console.error(`[ErrorLogger] ${newError.message}`, error);
    
    // Optional: Send to analytics or monitoring service
    // reportErrorToService(newError);
    
    return newError.id;
  };
  
  const clearErrors = () => setErrors([]);
  const removeError = (id: string) => setErrors(prev => prev.filter(e => e.id !== id));
  
  return { errors, logError, clearErrors, removeError };
}

export function ErrorBoundary({ children }: { children: React.ReactNode }) {
  const [hasError, setHasError] = useState(false);
  const [error, setError] = useState<Error | null>(null);
  const [componentStack, setComponentStack] = useState<string | null>(null);
  
  useEffect(() => {
    const handleGlobalError = (event: ErrorEvent) => {
      setHasError(true);
      setError(event.error || new Error(event.message));
      // Log to console for debugging
      console.error('[GlobalError]', event.error || event.message);
    };
    
    const handleRejection = (event: PromiseRejectionEvent) => {
      setHasError(true);
      setError(event.reason instanceof Error ? event.reason : new Error(String(event.reason)));
      console.error('[UnhandledRejection]', event.reason);
    };
    
    window.addEventListener('error', handleGlobalError);
    window.addEventListener('unhandledrejection', handleRejection);
    
    return () => {
      window.removeEventListener('error', handleGlobalError);
      window.removeEventListener('unhandledrejection', handleRejection);
    };
  }, []);
  
  if (hasError) {
    return (
      <div className="fixed inset-0 bg-red-50 flex items-center justify-center p-4 z-50">
        <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full p-6">
          <div className="flex items-center text-red-600 mb-4">
            <AlertCircle className="mr-2" />
            <h2 className="text-xl font-semibold">Application Error</h2>
          </div>
          
          <p className="mb-4">
            Something went wrong. Please try refreshing the page or contact support if the issue persists.
          </p>
          
          <div className="bg-gray-100 p-4 rounded mb-4 overflow-auto max-h-60">
            <p className="font-mono text-sm">{error?.message || 'Unknown error'}</p>
            {error?.stack && (
              <pre className="mt-2 text-xs overflow-auto">{error.stack}</pre>
            )}
            {componentStack && (
              <div className="mt-2 border-t border-gray-300 pt-2">
                <p className="font-semibold text-xs">Component Stack:</p>
                <pre className="text-xs overflow-auto">{componentStack}</pre>
              </div>
            )}
          </div>
          
          <div className="flex justify-end">
            <button 
              onClick={() => window.location.reload()}
              className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
            >
              Reload Application
            </button>
          </div>
        </div>
      </div>
    );
  }
  
  return <>{children}</>;
}

export function ErrorLogPanel() {
  const [isOpen, setIsOpen] = useState(false);
  const { errors, clearErrors, removeError } = useErrorLogger();
  const hasErrors = errors.length > 0;
  
  if (!hasErrors && !isOpen) return null;
  
  return (
    <div className="fixed bottom-0 right-0 z-50 w-full md:w-96 bg-white shadow-lg border-t border-l border-gray-200 rounded-tl-lg overflow-hidden">
      <div 
        className={`flex items-center justify-between p-2 ${hasErrors ? 'bg-red-100' : 'bg-gray-100'} cursor-pointer`}
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center">
          <AlertCircle className={`mr-2 ${hasErrors ? 'text-red-600' : 'text-gray-600'}`} size={18} />
          <span className="font-medium">
            {hasErrors ? `${errors.length} Error${errors.length > 1 ? 's' : ''}` : 'Error Log'}
          </span>
        </div>
        <div className="flex items-center">
          {hasErrors && (
            <button 
              onClick={(e) => { e.stopPropagation(); clearErrors(); }}
              className="mr-2 text-xs bg-gray-200 hover:bg-gray-300 px-2 py-1 rounded"
            >
              Clear
            </button>
          )}
          {isOpen ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
        </div>
      </div>
      
      {isOpen && (
        <div className="max-h-96 overflow-y-auto p-2">
          {errors.length === 0 ? (
            <p className="text-gray-500 text-center py-4">No errors logged</p>
          ) : (
            errors.map(error => (
              <ErrorLogEntry 
                key={error.id} 
                error={error} 
                onRemove={() => removeError(error.id)} 
              />
            ))
          )}
        </div>
      )}
    </div>
  );
}

function ErrorLogEntry({ error, onRemove }: { error: ErrorLogEntry, onRemove: () => void }) {
  const [expanded, setExpanded] = useState(false);
  
  const copyToClipboard = () => {
    const text = `
Error: ${error.message}
Timestamp: ${error.timestamp.toISOString()}
Source: ${error.source}
${error.stack ? `\nStack Trace:\n${error.stack}` : ''}
${error.componentStack ? `\nComponent Stack:\n${error.componentStack}` : ''}
    `.trim();
    
    navigator.clipboard.writeText(text);
  };
  
  return (
    <div className="mb-2 border border-gray-200 rounded overflow-hidden">
      <div className="flex items-center justify-between p-2 bg-gray-50">
        <div className="flex items-center">
          <button 
            onClick={() => setExpanded(!expanded)}
            className="mr-2 text-gray-600 hover:text-gray-800"
          >
            {expanded ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
          </button>
          <span className="font-medium truncate max-w-[200px]">{error.message}</span>
        </div>
        <div className="flex items-center">
          <button 
            onClick={copyToClipboard}
            className="mr-1 text-gray-600 hover:text-gray-800"
            title="Copy error details"
          >
            <Copy size={14} />
          </button>
          <button 
            onClick={onRemove}
            className="text-gray-600 hover:text-gray-800"
            title="Dismiss"
          >
            <X size={14} />
          </button>
        </div>
      </div>
      
      {expanded && (
        <div className="p-2 bg-gray-50 border-t border-gray-200">
          <div className="text-xs text-gray-500 mb-1">
            {error.timestamp.toLocaleString()}
            {error.source && <span className="ml-2 bg-gray-200 px-1 rounded">{error.source}</span>}
          </div>
          
          {error.stack && (
            <pre className="text-xs bg-gray-100 p-2 rounded overflow-x-auto mt-2">
              {error.stack}
            </pre>
          )}
          
          {error.componentStack && (
            <div className="mt-2">
              <div className="text-xs font-medium">Component Stack:</div>
              <pre className="text-xs bg-gray-100 p-2 rounded overflow-x-auto">
                {error.componentStack}
              </pre>
            </div>
          )}
        </div>
      )}
    </div>
  );
} 