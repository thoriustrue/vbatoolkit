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
    '\u00e2\u20ac\u201c': '\u2013', // en dash
    '\u00e2\u20ac\u201d': '\u2014', // em dash
    '\u00e2\u20ac\u02dc': '\u2018', // left single quote
    '\u00e2\u20ac\u2122': '\u2019', // right single quote
    '\u00e2\u20ac\u0153': '\u201c', // left double quote
    '\u00e2\u20ac\u009d': '\u201d', // right double quote
    '\u00e2\u20ac\u00a6': '\u2026', // ellipsis
    '\u00c2\u00a9': '\u00a9', // copyright
    '\u00c2\u00ae': '\u00ae', // registered trademark
    '\u00e2\u201e\u00a2': '\u2122', // trademark
  };
  
  // Apply all encoding fixes
  Object.entries(encodingFixes).forEach(([broken, replacement]) => {
    fixed = fixed.replace(new RegExp(broken, 'g'), replacement);
  });
  
  return fixed;
} 