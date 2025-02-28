import { useState } from 'react';
import { AlertCircle, X } from 'lucide-react';
import { LogEntry } from '../utils/types';

export function ErrorLog({ logs, onClear }: { 
  logs: LogEntry[];
  onClear: () => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const errorCount = logs.filter(l => l.type === 'error').length;

  return (
    <div className="fixed bottom-4 right-4 max-w-xs w-full bg-white dark:bg-gray-800 shadow-lg rounded-lg border border-red-200 dark:border-red-800">
      <div 
        className="flex items-center justify-between p-3 cursor-pointer"
        onClick={() => setExpanded(!expanded)}
      >
        <div className="flex items-center space-x-2">
          <AlertCircle className="text-red-500" size={18} />
          <span className="font-medium">{errorCount} error{errorCount !== 1 ? 's' : ''}</span>
        </div>
        <button
          onClick={(e) => {
            e.stopPropagation();
            onClear();
          }}
          className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
        >
          <X size={16} />
        </button>
      </div>
      
      {expanded && (
        <div className="max-h-64 overflow-y-auto p-3 border-t border-gray-100 dark:border-gray-700">
          {logs.map((log, index) => (
            <div 
              key={index}
              className={`text-sm p-2 rounded mb-1 ${
                log.type === 'error' ? 'bg-red-50 dark:bg-red-900/20 text-red-600 dark:text-red-400' :
                log.type === 'info' ? 'bg-blue-50 dark:bg-blue-900/20 text-blue-600 dark:text-blue-400' :
                'bg-green-50 dark:bg-green-900/20 text-green-600 dark:text-green-400'
              }`}
            >
              <div className="flex items-center space-x-2">
                <span className="flex-1">{log.message}</span>
                <span className="text-xs opacity-75">
                  {new Date(log.timestamp).toLocaleTimeString()}
                </span>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
} 