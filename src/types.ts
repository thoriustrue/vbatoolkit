// Log types
export type LogType = 'info' | 'error' | 'success' | 'warning';
export interface LogEntry {
  message: string;
  type: LogType;
}

// Callback types
export type LoggerCallback = (message: string, type: LogType) => void;
export type ProgressCallback = (progress: number) => void;

// Changelog types
export interface ChangelogChange {
  type: 'added' | 'fixed' | 'changed' | 'removed';
  description: string;
}

export interface ChangelogEntry {
  version: string;
  date: string;
  changes: ChangelogChange[];
}

// Error types
export interface ErrorLogEntry {
  id: string;
  timestamp: Date;
  message: string;
  stack?: string;
  componentStack?: string;
  source?: string;
} 