import React from 'react';
import { Info, AlertCircle, CheckCircle, X } from 'lucide-react';

interface Log {
  message: string;
  type: 'info' | 'error' | 'success';
}

interface LogViewerProps {
  logs: Log[];
  onClearLogs: () => void;
}

export function LogViewer({ logs, onClearLogs }: LogViewerProps) {
  const getLogIcon = (type: 'info' | 'error' | 'success') => {
    switch (type) {
      case 'info': return <Info className="w-4 h-4 text-blue-500" />;
      case 'error': return <AlertCircle className="w-4 h-4 text-red-500" />;
      case 'success': return <CheckCircle className="w-4 h-4 text-green-500" />;
    }
  };

  if (logs.length === 0) {
    return (
      <div className="text-center py-4 text-gray-500 text-sm">
        No logs to display
      </div>
    );
  }

  return (
    <div className="mt-4 border rounded-md overflow-hidden">
      <div className="bg-gray-100 px-4 py-2 flex justify-between items-center">
        <h3 className="text-sm font-medium">Process Logs</h3>
        <button 
          onClick={onClearLogs}
          className="text-gray-500 hover:text-gray-700"
          title="Clear logs"
        >
          <X size={16} />
        </button>
      </div>
      <div className="max-h-60 overflow-y-auto p-2">
        {logs.map((log, index) => (
          <div 
            key={index} 
            className={`flex items-start p-2 text-sm ${
              index % 2 === 0 ? 'bg-gray-50' : 'bg-white'
            }`}
          >
            <div className="flex-shrink-0 mt-0.5 mr-2">
              {getLogIcon(log.type)}
            </div>
            <div className="flex-1">
              <p className={`
                ${log.type === 'error' ? 'text-red-700' : ''}
                ${log.type === 'success' ? 'text-green-700' : ''}
                ${log.type === 'info' ? 'text-gray-700' : ''}
              `}>
                {log.message}
              </p>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
} 