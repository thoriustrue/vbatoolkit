import { useState, useEffect } from 'react';
import { Clock, ChevronDown, ChevronUp, Tag } from 'lucide-react';

interface ChangelogEntry {
  version: string;
  date: string;
  changes: {
    type: 'added' | 'changed' | 'fixed' | 'removed';
    description: string;
  }[];
}

// This would typically be loaded from a JSON file or API
const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '1.0.0',
    date: '2025-02-28',
    changes: [
      { type: 'added', description: 'Initial release of VBA Toolkit' },
      { type: 'added', description: 'VBA password removal functionality' },
      { type: 'added', description: 'Excel security settings removal' }
    ]
  },
  {
    version: '1.0.1',
    date: '2025-02-29',
    changes: [
      { type: 'fixed', description: 'Fixed file corruption issues with large Excel files' },
      { type: 'added', description: 'Added CRC validation for ZIP archives' },
      { type: 'changed', description: 'Improved error handling and logging' }
    ]
  },
  {
    version: '1.1.0',
    date: '2025-03-01',
    changes: [
      { type: 'added', description: 'Added technical error logging panel' },
      { type: 'added', description: 'Added changelog viewer' },
      { type: 'fixed', description: 'Fixed dependency issues with adm-zip' },
      { type: 'changed', description: 'Updated build process for better reliability' }
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