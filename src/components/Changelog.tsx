import React, { useState } from 'react';
import { ChevronDown, ChevronUp } from 'lucide-react';

interface ChangelogEntry {
  version: string;
  date: string;
  changes: {
    type: 'added' | 'fixed' | 'changed' | 'removed';
    description: string;
  }[];
}

// Define the changelog data directly in this component
const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '0.1.1',
    date: '2024-06-20',
    changes: [
      { type: 'added', description: 'Added VBA password removal functionality' },
      { type: 'added', description: 'Implemented macro auto-enable features' },
      { type: 'added', description: 'Added error logging system' }
    ]
  },
  {
    version: '0.1.0',
    date: '2024-06-15',
    changes: [
      { type: 'added', description: 'Initial release of VBA Toolkit' },
      { type: 'added', description: 'Basic file handling capabilities' },
      { type: 'added', description: 'User interface for file operations' }
    ]
  }
];

export function Changelog() {
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
      <div className="ml-6 mt-2 space-y-2 max-h-80 overflow-y-auto">
        {CHANGELOG_DATA.map((entry, index) => (
          <div key={index} className="space-y-2 mb-4">
            <h3 className="font-medium">v{entry.version} - {entry.date}</h3>
            <ul className="list-disc pl-4 space-y-1">
              {entry.changes.map((change, i) => (
                <li key={i} className={
                  change.type === 'added' ? 'text-green-600 dark:text-green-400' :
                  change.type === 'fixed' ? 'text-amber-600 dark:text-amber-400' :
                  'text-blue-600 dark:text-blue-400'
                }>
                  {change.description}
                </li>
              ))}
            </ul>
          </div>
        ))}
      </div>
    </details>
  );
} 