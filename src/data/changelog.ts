// Define the types for our changelog data
export interface ChangelogEntry {
  version: string;
  date: string;
  changes: {
    type: 'added' | 'fixed' | 'changed' | 'removed';
    description: string;
  }[];
}

// Export the changelog data
export const CHANGELOG_DATA: ChangelogEntry[] = [
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