import { useState } from 'react';
import { ChevronDown, ChevronUp } from 'lucide-react';
import { ChangelogEntry } from '../types';

export const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: "1.1.1",
    date: "2025-02-28",
    changes: [
      { type: "added", description: "Added automated changelog update script" },
      { type: "added", description: "Added pre-commit hook for changelog reminders" },
      { type: "fixed", description: "Fixed ES module compatibility issues" }
    ]
  },
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

interface ChangelogProps {
  children?: React.ReactNode;
}

/**
 * Changelog component that displays version history in a collapsible section
 * @param children Optional content to display inside the changelog
 */
export function Changelog({ children }: ChangelogProps) {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <details 
      className="mt-4 text-sm text-gray-600 dark:text-gray-300"
      open={isOpen}
    >
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

/**
 * Renders a formatted changelog entry
 * @param entry The changelog entry to render
 */
export function ChangelogEntryComponent({ entry }: { entry: ChangelogEntry }) {
  return (
    <div className="mb-4">
      <h4 className="font-semibold">
        Version {entry.version} <span className="text-gray-500 font-normal">({entry.date})</span>
      </h4>
      <ul className="mt-1 space-y-1">
        {entry.changes.map((change: { type: string; description: string }, idx: number) => (
          <li key={idx} className="text-sm">
            <span className={`
              inline-block px-2 py-0.5 rounded text-xs mr-2
              ${change.type === 'added' ? 'bg-green-100 text-green-800' : ''}
              ${change.type === 'fixed' ? 'bg-blue-100 text-blue-800' : ''}
              ${change.type === 'changed' ? 'bg-yellow-100 text-yellow-800' : ''}
              ${change.type === 'removed' ? 'bg-red-100 text-red-800' : ''}
            `}>
              {change.type}
            </span>
            {change.description}
          </li>
        ))}
      </ul>
    </div>
  );
} 