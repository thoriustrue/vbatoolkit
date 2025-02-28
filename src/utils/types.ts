export interface LogEntry {
  message: string;
  type: 'error' | 'info' | 'success';
  timestamp: number;
} 