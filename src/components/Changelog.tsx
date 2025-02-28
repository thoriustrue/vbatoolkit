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

export function Changelog() {
  const [isOpen, setIsOpen] = useState(false);
  const [expandedVersions, setExpandedVersions] = useState<string[]>([CHANGELOG_DATA[0]?.version || '']);
  
  const toggleVersion = (version: string) => {
    setExpandedVersions(prev => 
      prev.includes(version) 
        ? prev.filter(v => v !== version) 
        : [...prev, version]
    );
  };
  
  return (
    <div className="bg-white rounded-lg shadow-md overflow-hidden">
      <div 
        className="flex items-center justify-between p-4 bg-gray-50 cursor-pointer"
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center">
          <Clock className="mr-2 text-gray-600" size={18} />
          <h2 className="text-lg font-medium">Changelog</h2>
        </div>
        {isOpen ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
      </div>
      
      {isOpen && (
        <div className="p-4 max-h-96 overflow-y-auto">
          {CHANGELOG_DATA.map(entry => (
            <div key={entry.version} className="mb-4 last:mb-0">
              <div 
                className="flex items-center justify-between cursor-pointer"
                onClick={() => toggleVersion(entry.version)}
              >
                <div className="flex items-center">
                  <Tag className="mr-2 text-blue-600" size={16} />
                  <h3 className="text-md font-medium">Version {entry.version}</h3>
                  <span className="ml-2 text-sm text-gray-500">{entry.date}</span>
                </div>
                {expandedVersions.includes(entry.version) ? 
                  <ChevronUp size={16} /> : 
                  <ChevronDown size={16} />
                }
              </div>
              
              {expandedVersions.includes(entry.version) && (
                <div className="mt-2 pl-6 border-l-2 border-gray-200">
                  <ul className="space-y-2">
                    {entry.changes.map((change, idx) => (
                      <li key={idx} className="flex">
                        <span className={`
                          inline-block w-16 text-xs font-medium rounded px-2 py-1 mr-2 text-center
                          ${change.type === 'added' ? 'bg-green-100 text-green-800' : ''}
                          ${change.type === 'changed' ? 'bg-blue-100 text-blue-800' : ''}
                          ${change.type === 'fixed' ? 'bg-yellow-100 text-yellow-800' : ''}
                          ${change.type === 'removed' ? 'bg-red-100 text-red-800' : ''}
                        `}>
                          {change.type}
                        </span>
                        <span>{change.description}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
} 