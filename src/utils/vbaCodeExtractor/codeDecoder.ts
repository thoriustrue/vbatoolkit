/**
 * Cleans and decodes VBA code from various encodings
 * @param code The raw VBA code to clean and decode
 * @returns The cleaned and decoded VBA code
 */
export function cleanAndDecodeVBACode(code: string): string {
  // Remove common binary artifacts and control characters
  let cleanedCode = code.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  
  // Try to detect and fix encoding issues
  cleanedCode = fixEncodingIssues(cleanedCode);
  
  // Remove trailing nulls and whitespace
  cleanedCode = cleanedCode.replace(/\x00+$/, '').trim();
  
  // Normalize line endings
  cleanedCode = cleanedCode.replace(/\r\n|\r|\n/g, '\r\n');
  
  return cleanedCode;
}

/**
 * Attempts to fix common encoding issues in VBA code
 * @param code The code to fix
 * @returns The fixed code
 */
function fixEncodingIssues(code: string): string {
  // Fix UTF-16 artifacts (common in Excel VBA)
  let fixed = code.replace(/\u0000/g, '');
  
  // Fix common Unicode replacement characters
  fixed = fixed.replace(/\uFFFD/g, '?');
  
  // Fix common encoding issues with special characters
  const encodingFixes: Record<string, string> = {
    'â€"': '–', // en dash
    'â€"': '—', // em dash
    'â€˜': ''', // left single quote
    'â€™': ''', // right single quote
    'â€œ': '"', // left double quote
    'â€': '"', // right double quote
    'â€¦': '…', // ellipsis
    'Â©': '©', // copyright
    'Â®': '®', // registered trademark
    'â„¢': '™', // trademark
  };
  
  // Apply all encoding fixes
  Object.entries(encodingFixes).forEach(([broken, fixed]) => {
    fixed = fixed.replace(new RegExp(broken, 'g'), fixed);
  });
  
  return fixed;
} 