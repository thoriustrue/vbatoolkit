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

// Define the changelog data to match CHANGELOG.md
const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '1.1.0',
    date: '2025-03-01',
    changes: [
      { type: 'added', description: 'Technical error logging panel for debugging' },
      { type: 'added', description: 'Changelog viewer to track version history' },
      { type: 'added', description: 'Improved error handling with detailed stack traces' },
      { type: 'fixed', description: 'Dependency issues with adm-zip' },
      { type: 'fixed', description: 'Build process reliability improvements' },
      { type: 'fixed', description: 'File corruption when processing large Excel files' },
      { type: 'changed', description: 'Updated build configuration for GitHub Pages' },
      { type: 'changed', description: 'Improved ZIP validation with CRC checks' },
      { type: 'added', description: 'Enhanced VBA code extraction with multiple encoding support' },
      { type: 'fixed', description: 'Overly aggressive file integrity fixes that removed necessary files' },
      { type: 'added', description: 'Maximum trust settings to prevent Excel security prompts' },
      { type: 'changed', description: 'Simplified UI by removing alternate method tab' }
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
      <div className="ml-6 mt-2 space-y-4 max-h-80 overflow-y-auto">
        {CHANGELOG_DATA.map((entry, index) => (
          <div key={index} className="space-y-3 mb-4">
            <h3 className="font-medium text-base">v{entry.version} - {entry.date}</h3>
            
            {/* Group changes by type */}
            {['added', 'fixed', 'changed', 'removed'].map(type => {
              const typeChanges = entry.changes.filter(change => change.type === type);
              if (typeChanges.length === 0) return null;
              
              return (
                <div key={type} className="space-y-1">
                  <h4 className="font-medium capitalize">{type}:</h4>
                  <ul className="list-disc pl-5 space-y-1">
                    {typeChanges.map((change, i) => (
                      <li key={i} className={
                        type === 'added' ? 'text-green-600 dark:text-green-400' :
                        type === 'fixed' ? 'text-amber-600 dark:text-amber-400' :
                        type === 'changed' ? 'text-blue-600 dark:text-blue-400' :
                        'text-red-600 dark:text-red-400'
                      }>
                        {change.description}
                      </li>
                    ))}
                  </ul>
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </details>
  );
} 