import { LoggerCallback } from '../../types';

/**
 * Cleans and decodes VBA code
 * @param code The raw VBA code to clean
 * @returns The cleaned and decoded VBA code
 */
export function cleanAndDecodeVBACode(code: string): string {
  if (!code || code.trim() === '') {
    return '\'No code content available';
  }
  
  try {
    // Remove common binary artifacts
    let cleanedCode = code
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control characters
      .replace(/\r\n/g, '\n')                           // Normalize line endings
      .replace(/\r/g, '\n')                             // Convert remaining CR to LF
      .replace(/\n{3,}/g, '\n\n');                      // Remove excessive blank lines
    
    // Fix common encoding issues
    cleanedCode = fixEncodingIssues(cleanedCode);
    
    // Add proper line endings if missing
    if (!cleanedCode.endsWith('\n')) {
      cleanedCode += '\n';
    }
    
    return cleanedCode;
  } catch (error) {
    console.error('Error cleaning VBA code:', error);
    return code; // Return original if cleaning fails
  }
}

/**
 * Fixes common encoding issues in VBA code
 * @param code The code to fix
 * @returns The fixed code
 */
function fixEncodingIssues(code: string): string {
  // Replace common encoding artifacts
  return code
    .replace(/Ã¢â‚¬â„¢/g, "'")     // Smart single quote
    .replace(/Ã¢â‚¬Å"/g, '"')      // Smart double quote open
    .replace(/Ã¢â‚¬Â/g, '"')       // Smart double quote close
    .replace(/Ã¢â‚¬â€œ/g, '-')      // Em dash
    .replace(/Ã¢â‚¬â€/g, '-')       // En dash
    .replace(/Ã¢â‚¬Â¦/g, '...')     // Ellipsis
    .replace(/Ã‚/g, '')            // Non-breaking space artifact
    .replace(/Ã¯Â¿Â½/g, '?');       // Replacement character
}

/**
 * Attempts to extract meaningful code from a corrupted VBA module
 * @param corruptedCode The corrupted VBA code
 * @param logger Optional logger callback
 * @returns The best attempt at extracting meaningful code
 */
export function extractFromCorruptedCode(
  corruptedCode: string, 
  logger?: LoggerCallback
): string {
  if (!corruptedCode || corruptedCode.trim() === '') {
    return '\'No code content available';
  }
  
  try {
    // Log the attempt if logger is provided
    if (logger) {
      logger('Attempting to extract code from corrupted module', 'info');
    }
    
    // Look for VBA code patterns
    const codePatterns = [
      // Sub/Function declarations
      /(?:Public |Private |Friend )?(?:Sub|Function|Property Get|Property Let|Property Set)[^\n]+\n[\s\S]+?End (?:Sub|Function|Property)/gi,
      
      // Variable declarations
      /(?:Dim|Public|Private|Global|Static|Const) [^\n]+/gi,
      
      // Type declarations
      /Type[\s\S]+?End Type/gi,
      
      // Enum declarations
      /Enum[\s\S]+?End Enum/gi
    ];
    
    let extractedParts: string[] = [];
    
    // Apply each pattern and collect matches
    for (const pattern of codePatterns) {
      const matches = corruptedCode.match(pattern);
      if (matches) {
        extractedParts = extractedParts.concat(matches);
      }
    }
    
    if (extractedParts.length > 0) {
      // Join the extracted parts with blank lines
      let result = extractedParts.join('\n\n');
      
      // Clean the result
      result = cleanAndDecodeVBACode(result);
      
      if (logger) {
        logger(`Successfully extracted ${extractedParts.length} code fragments`, 'success');
      }
      
      return result;
    } else {
      // If no patterns matched, return a placeholder
      return '\'Code could not be extracted from corrupted module';
    }
  } catch (error) {
    if (logger) {
      logger(`Error extracting from corrupted code: ${error instanceof Error ? error.message : String(error)}`, 'error');
    }
    return '\'Error extracting code from corrupted module';
  }
} 