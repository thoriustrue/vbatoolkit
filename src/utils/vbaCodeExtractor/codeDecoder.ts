import { LoggerCallback } from '../../types';

/**
 * Cleans and decodes raw VBA code
 * @param code The raw VBA code to clean
 * @returns Cleaned and decoded VBA code
 */
export function cleanAndDecodeVBACode(code: string): string {
  if (!code) return '';
  
  // Remove binary artifacts and control characters
  let cleanedCode = code.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  
  // Remove VBA attributes section
  const attributeEndIndex = cleanedCode.search(/(?:^|\r\n)(?!Attribute VB_)/m);
  if (attributeEndIndex > 0) {
    cleanedCode = cleanedCode.substring(attributeEndIndex).trim();
  }
  
  // Fix common encoding issues
  cleanedCode = cleanedCode
    .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'") // Replace various quote marks with apostrophe
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"') // Replace various quote marks with double quote
    .replace(/[\u2013\u2014\u2015]/g, '-') // Replace various dashes with hyphen
    .replace(/[\uFFFD\uFFFE\uFFFF]/g, '') // Remove Unicode replacement characters
    .replace(/\r\n/g, '\n') // Normalize line endings
    .replace(/\r/g, '\n') // Convert remaining CR to LF
    .replace(/\n\n\n+/g, '\n\n'); // Remove excessive blank lines
  
  return cleanedCode;
}

/**
 * Attempts to extract meaningful code from corrupted VBA modules
 * @param corruptedCode The corrupted VBA code
 * @param moduleName The name of the module
 * @returns Extracted code or placeholder message
 */
export function extractFromCorruptedCode(corruptedCode: string, moduleName: string): string {
  if (!corruptedCode) {
    return `' Code could not be extracted for module: ${moduleName}`;
  }
  
  // Try to find code sections using common patterns
  let extractedCode = '';
  
  // Pattern 1: Look for Sub/Function declarations
  const procedureMatches = corruptedCode.match(/(?:Public |Private |Friend )?(?:Sub|Function|Property Get|Property Let|Property Set)[^\n]+(?:\n(?!\s*(?:Sub|Function|Property|End Sub|End Function|End Property))[^\n]+)*/g);
  
  if (procedureMatches && procedureMatches.length > 0) {
    extractedCode = procedureMatches.join('\n\n');
    return cleanAndDecodeVBACode(extractedCode);
  }
  
  // Pattern 2: Look for Option statements
  const optionMatches = corruptedCode.match(/Option [^\n]+/g);
  if (optionMatches && optionMatches.length > 0) {
    extractedCode = optionMatches.join('\n');
    
    // Try to find variable declarations after options
    const declarationMatches = corruptedCode.match(/(?:Public |Private |Dim |Const )[^\n]+/g);
    if (declarationMatches && declarationMatches.length > 0) {
      extractedCode += '\n\n' + declarationMatches.join('\n');
    }
    
    return cleanAndDecodeVBACode(extractedCode);
  }
  
  // Pattern 3: Look for comments
  const commentMatches = corruptedCode.match(/'[^\n]+/g);
  if (commentMatches && commentMatches.length > 0) {
    extractedCode = commentMatches.join('\n');
    return cleanAndDecodeVBACode(extractedCode);
  }
  
  // If all else fails, return a placeholder with the original corrupted code
  return `' Code could not be properly extracted for module: ${moduleName}\n' Partial content:\n\n${cleanAndDecodeVBACode(corruptedCode.substring(0, 500))}`;
} 