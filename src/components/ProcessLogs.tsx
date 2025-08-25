import React from 'react';
import { createProcessLogsFile } from '../utils/vbaCodeExtractor';

interface ProcessLogsProps {
  logs: string[];
  processType: string;
}

const ProcessLogs: React.FC<ProcessLogsProps> = ({ logs, processType }) => {
  const handleDownload = () => {
    const logsBlob = createProcessLogsFile(logs, processType);
    const url = URL.createObjectURL(logsBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${processType.toLowerCase().replace(/\s+/g, '_')}_logs.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="bg-gray-50 p-4 mt-4 rounded-lg border-l-4 border-blue-500">
      <div className="flex justify-between items-center mb-2">
        <h3 className="text-lg font-semibold">{processType} Logs</h3>
        <button 
          onClick={handleDownload}
          className="px-3 py-1 bg-blue-600 text-white rounded text-sm hover:bg-blue-700"
        >
          Download Logs
        </button>
      </div>
      <div className="max-h-64 overflow-auto font-mono text-sm bg-white p-2 rounded border">
        {logs.map((log, index) => (
          <div key={index} className="py-1 border-b border-gray-100 last:border-b-0">
            {log}
          </div>
        ))}
      </div>
    </div>
  );
};

export default ProcessLogs; 