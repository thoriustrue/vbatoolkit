#!/usr/bin/env node
/**
 * Script to update the Changelog.tsx file with new entries
 * 
 * Usage:
 *   node scripts/update-changelog.js --version "1.2.0" [options]
 * 
 * Options:
 *   --version: Required. Version number for the new entry
 *   --add: Description of a new feature (can be used multiple times)
 *   --fix: Description of a bug fix (can be used multiple times)
 *   --change: Description of a change (can be used multiple times)
 *   --remove: Description of a removal (can be used multiple times)
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory name
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// If no arguments are provided, show usage example
if (process.argv.length <= 2) {
  console.log(`
Usage Examples:
  
  # Add a new version with multiple types of changes
  node scripts/update-changelog.js --version "1.2.0" --add "New feature A" --add "New feature B" --fix "Fixed bug X" --change "Updated dependency Y"
  
  # Add a version with just one type of change
  node scripts/update-changelog.js --version "1.1.1" --fix "Critical security fix"
  
  # Using with npm script
  npm run changelog -- --version "1.2.0" --add "New feature" --fix "Bug fix"
`);
  process.exit(0);
}

// Parse command line arguments
const args = process.argv.slice(2);
const options = {
  version: null,
  add: [],
  fix: [],
  change: [],
  remove: []
};

for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  
  if (arg === '--version' && i + 1 < args.length) {
    options.version = args[++i];
  } else if (arg === '--add' && i + 1 < args.length) {
    options.add.push(args[++i]);
  } else if (arg === '--fix' && i + 1 < args.length) {
    options.fix.push(args[++i]);
  } else if (arg === '--change' && i + 1 < args.length) {
    options.change.push(args[++i]);
  } else if (arg === '--remove' && i + 1 < args.length) {
    options.remove.push(args[++i]);
  }
}

// Validate required arguments
if (!options.version) {
  console.error('Error: --version is required');
  process.exit(1);
}

// Validate that at least one change is provided
if (options.add.length === 0 && options.fix.length === 0 && 
    options.change.length === 0 && options.remove.length === 0) {
  console.error('Error: At least one change (--add, --fix, --change, or --remove) is required');
  process.exit(1);
}

// Path to the Changelog.tsx file
const changelogPath = path.join(__dirname, '..', 'src', 'components', 'Changelog.tsx');

try {
  // Read the current changelog file
  const changelogContent = fs.readFileSync(changelogPath, 'utf8');
  
  // Find the CHANGELOG_DATA array
  const match = changelogContent.match(/export const CHANGELOG_DATA: ChangelogEntry\[\] = \[([\s\S]*?)\];/);
  
  if (!match) {
    console.error('Error: Could not find CHANGELOG_DATA array in Changelog.tsx');
    process.exit(1);
  }
  
  // Construct the new entry
  let newEntry = `  {
    version: "${options.version}",
    date: "${new Date().toISOString().split('T')[0]}",
    changes: [`;
  
  // Add new features
  for (const item of options.add) {
    newEntry += `\n      { type: "added", description: "${item}" },`;
  }
  
  // Add bug fixes
  for (const item of options.fix) {
    newEntry += `\n      { type: "fixed", description: "${item}" },`;
  }
  
  // Add changes
  for (const item of options.change) {
    newEntry += `\n      { type: "changed", description: "${item}" },`;
  }
  
  // Add removals
  for (const item of options.remove) {
    newEntry += `\n      { type: "removed", description: "${item}" },`;
  }
  
  // Remove trailing comma if needed
  if (newEntry.endsWith(',')) {
    newEntry = newEntry.slice(0, -1);
  }
  
  newEntry += `\n    ]\n  }`;
  
  // Insert the new entry at the beginning of the array
  const updatedContent = changelogContent.replace(
    /export const CHANGELOG_DATA: ChangelogEntry\[\] = \[/,
    `export const CHANGELOG_DATA: ChangelogEntry[] = [\n${newEntry},`
  );
  
  // Write the updated content back to the file
  fs.writeFileSync(changelogPath, updatedContent, 'utf8');
  
  console.log(`Successfully updated Changelog.tsx with version ${options.version}`);
  
} catch (error) {
  console.error('Error updating changelog:', error.message);
  process.exit(1);
} 