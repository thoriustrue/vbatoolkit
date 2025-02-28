import React, { useState } from 'react';
import { Box, Button, Typography, Paper, Divider } from '@mui/material';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import DownloadIcon from '@mui/icons-material/Download';
import { createProcessLogsFile } from '../utils/vbaCodeExtractor';

interface ProcessLogsProps {
  logs: string[];
  processType: string;
}

const ProcessLogs: React.FC<ProcessLogsProps> = ({ logs, processType }) => {
  const [copied, setCopied] = useState(false);

  const handleCopy = () => {
    const logsText = logs.join('\n');
    navigator.clipboard.writeText(logsText).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

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
    <Paper 
      elevation={3} 
      sx={{ 
        p: 2, 
        mt: 2, 
        mb: 2, 
        maxHeight: '300px', 
        overflow: 'auto',
        backgroundColor: '#f5f5f5',
        borderLeft: '4px solid #2196f3'
      }}
    >
      <Box display="flex" justifyContent="space-between" alignItems="center" mb={1}>
        <Typography variant="h6" component="h3">
          {processType} Logs
        </Typography>
        <Box>
          <Button 
            startIcon={<ContentCopyIcon />} 
            onClick={handleCopy} 
            color={copied ? "success" : "primary"}
            size="small"
            sx={{ mr: 1 }}
          >
            {copied ? "Copied!" : "Copy Logs"}
          </Button>
          <Button 
            startIcon={<DownloadIcon />} 
            onClick={handleDownload} 
            color="primary"
            size="small"
          >
            Download
          </Button>
        </Box>
      </Box>
      <Divider sx={{ mb: 1 }} />
      <Box sx={{ fontFamily: 'monospace', whiteSpace: 'pre-wrap', fontSize: '0.85rem' }}>
        {logs.map((log, index) => (
          <Typography 
            key={index} 
            variant="body2" 
            component="div" 
            sx={{ 
              py: 0.5,
              borderBottom: index < logs.length - 1 ? '1px dashed rgba(0,0,0,0.1)' : 'none'
            }}
          >
            {log}
          </Typography>
        ))}
      </Box>
    </Paper>
  );
};

export default ProcessLogs; 