import { useState, useEffect } from 'react';
import { Clock, ChevronDown, ChevronUp, Tag } from 'lucide-react';
import { ChangelogEntry } from '../types';

export const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '1.1.0',
    date: '2025-03-01',
    changes: [
      { type: 'added', description: 'Technical error logging panel for debugging' },
      { type: 'added', description: 'Changelog viewer to track version history' },
      { type: 'added', description: 'Improved error handling with detailed stack traces' },
      { type: 'added', description: 'Enhanced VBA code extraction with multiple encoding support' },
      { type: 'added', description: 'Maximum trust settings to prevent Excel security prompts' },
      { type: 'fixed', description: 'Dependency issues with adm-zip' },
      { type: 'fixed', description: 'Build process reliability improvements' },
      { type: 'fixed', description: 'File corruption when processing large Excel files' },
      { type: 'fixed', description: 'Overly aggressive file integrity fixes that removed necessary files' },
      { type: 'fixed', description: 'Excel repair dialog issues when opening unprotected files' },
      { type: 'changed', description: 'Updated build configuration for GitHub Pages' },
      { type: 'changed', description: 'Improved ZIP validation with CRC checks' },
      { type: 'changed', description: 'Simplified UI by removing alternate method tab' },
      { type: 'changed', description: 'More robust VBA project binary parsing' }
    ]
  },
  {
    version: '1.0.1',
    date: '2025-02-29',
    changes: [
      { type: 'fixed', description: 'File corruption issues with large Excel files' },
      { type: 'fixed', description: 'Missing dependencies in package.json' },
      { type: 'added', description: 'CRC validation for ZIP archives' },
      { type: 'added', description: 'Better error messages for common failures' },
      { type: 'changed', description: 'Improved error handling and logging' },
      { type: 'changed', description: 'Updated dependency management' }
    ]
  },
  {
    version: '1.0.0',
    date: '2025-02-28',
    changes: [
      { type: 'added', description: 'Initial release of VBA Toolkit' },
      { type: 'added', description: 'VBA password removal functionality' },
      { type: 'added', description: 'Excel security settings removal' },
      { type: 'added', description: 'Support for .xlsm, .xls, and .xlsb files' }
    ]
  }
];

export function Changelog({ children }: { children: React.ReactNode }) {
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