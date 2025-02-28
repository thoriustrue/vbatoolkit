#!/usr/bin/env node

/**
 * Pre-commit hook to remind developers to update the changelog
 * 
 * Installation:
 * 1. Make this script executable: chmod +x scripts/pre-commit.js
 * 2. Create a symlink in the git hooks directory:
 *    ln -s ../../scripts/pre-commit.js .git/hooks/pre-commit
 * 
 * This script will:
 * 1. Check if any significant files have been modified
 * 2. Check if the changelog has been updated
 * 3. If significant changes are detected but the changelog hasn't been updated,
 *    prompt the user to confirm whether to proceed with the commit
 */

import { execSync } from 'child_process';
import readline from 'readline';
import { fileURLToPath } from 'url';
import path from 'path';

// Get the directory name
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// If run directly (not as a hook), show installation instructions
if (import.meta.url === `file://${process.argv[1]}` && !process.env.GIT_HOOK_RUNNING) {
  console.log(`
This script is intended to be used as a Git pre-commit hook.

Installation:
1. Make this script executable:
   chmod +x scripts/pre-commit.js

2. Create a symlink in the git hooks directory:
   ln -s ../../scripts/pre-commit.js .git/hooks/pre-commit

Once installed, this hook will remind you to update the changelog when committing significant changes.
  `);
  process.exit(0);
}

// Get the list of staged files
function getStagedFiles() {
  try {
    // Set environment variable to indicate we're running as a hook
    process.env.GIT_HOOK_RUNNING = 'true';
    
    const output = execSync('git diff --cached --name-only').toString().trim();
    return output ? output.split('\n') : [];
  } catch (error) {
    console.error('Error getting staged files:', error.message);
    return [];
  }
}

// Check if any significant files have been modified
const significantPatterns = [
  /^src\/.*\.(ts|tsx|js|jsx)$/, // Source code
  /^vite\.config\.ts$/,         // Build configuration
  /^package\.json$/              // Dependencies
];

const stagedFiles = getStagedFiles();

const significantChanges = stagedFiles.some(file => 
  significantPatterns.some(pattern => pattern.test(file))
);

// Check if the changelog has been updated
const changelogUpdated = stagedFiles.includes('src/components/Changelog.tsx');

// If significant changes are detected but the changelog hasn't been updated,
// prompt the user to confirm whether to proceed with the commit
if (significantChanges && !changelogUpdated) {
  console.log('\x1b[33m%s\x1b[0m', '⚠️  Warning: You have made significant changes but have not updated the changelog.');
  console.log('Consider updating the changelog with:');
  console.log('\x1b[36m%s\x1b[0m', '  npm run changelog -- --version "x.y.z" --add "Your new feature" --fix "Your bug fix" --change "Your change"');
  
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  rl.question('\x1b[33m%s\x1b[0m', 'Do you want to continue with the commit anyway? (y/N): ', (answer) => {
    rl.close();
    
    if (answer.toLowerCase() !== 'y') {
      console.log('Commit aborted. Please update the changelog before committing.');
      process.exit(1);
    }
    
    console.log('Proceeding with commit without updating the changelog.');
    process.exit(0);
  });
} else {
  // No significant changes or changelog has been updated, proceed with commit
  process.exit(0);
} 