import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';

/**
 * Enhanced VBA code extraction with proper binary parsing
 */

interface VBADirectoryEntry {
  name: string;
  type: VBAModuleType;
  offset: number;
  size: number;
}

/**
 * Parse VBA directory to find all modules
 */
function parseVBADirectory(vbaData: Uint8Array, logger: LoggerCallback): VBADirectoryEntry[] {
  const entries: VBADirectoryEntry[] = [];
  
  try {
    // Convert to string for pattern matching (this is a simplified approach)
    const vbaText = new TextDecoder('latin1', { fatal: false }).decode(vbaData);
    
    // Look for directory entries in various formats
    const patterns = [
      // Standard module entries
      /Module=([^\r\n]+)/g,
      /Class=([^\r\n]+)/g,
      /Document=([^\r\n]+)/g,
      // Alternative patterns
      /VB_Name\s*=\s*"([^"]+)"/g,
      /Attribute\s+VB_Name\s*=\s*"([^"]+)"/g
    ];
    
    const moduleTypes = [VBAModuleType.Standard, VBAModuleType.Class, VBAModuleType.Document, VBAModuleType.Unknown, VBAModuleType.Unknown];
    
    for (let patternIndex = 0; patternIndex < patterns.length; patternIndex++) {
      const pattern = patterns[patternIndex];
      const moduleType = moduleTypes[patternIndex];
      let match;
      
      while ((match = pattern.exec(vbaText)) !== null) {
        const name = match[1].trim();
        if (name && !entries.some(e => e.name === name)) {
          entries.push({
            name,
            type: moduleType,
            offset: match.index,
            size: 0 // Will be calculated later
          });
        }
      }
    }
    
    logger(`Found ${entries.length} VBA modules in directory`, 'info');
    
  } catch (error) {
    logger(`Error parsing VBA directory: ${error instanceof Error ? error.message : String(error)}`, 'warning');
  }
  
  return entries;
}

/**
 * Extract actual VBA code for a module
 */
function extractModuleCode(vbaData: Uint8Array, moduleName: string, logger: LoggerCallback): string {
  try {
    // Convert to text for searching
    const vbaText = new TextDecoder('latin1', { fatal: false }).decode(vbaData);
    
    // Find the module's attribute section
    const namePattern = new RegExp(`Attribute\\s+VB_Name\\s*=\\s*"${moduleName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i');
    const match = namePattern.exec(vbaText);
    
    if (!match) {
      return `' Module ${moduleName} found in directory but code could not be extracted`;
    }
    
    const startIndex = match.index;
    
    // Find the end of this module (start of next module or end of file)
    const nextModulePattern = /Attribute\s+VB_Name\s*=/g;
    nextModulePattern.lastIndex = startIndex + match[0].length;
    const nextMatch = nextModulePattern.exec(vbaText);
    const endIndex = nextMatch ? nextMatch.index : vbaText.length;
    
    // Extract the module content
    let moduleContent = vbaText.substring(startIndex, endIndex);
    
    // Clean up the code
    moduleContent = cleanVBACode(moduleContent);
    
    if (moduleContent.trim().length === 0) {
      return `' Module ${moduleName} appears to be empty or corrupted`;
    }
    
    return moduleContent;
    
  } catch (error) {
    logger(`Error extracting code for module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`, 'warning');
    return `' Error extracting code for module ${moduleName}`;
  }
}

/**
 * Clean and format VBA code
 */
function cleanVBACode(rawCode: string): string {
  // Remove binary artifacts and control characters
  // eslint-disable-next-line no-control-regex
  let cleanCode = rawCode.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  
  // Normalize line endings
  cleanCode = cleanCode.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  // Remove excessive blank lines
  cleanCode = cleanCode.replace(/\n\n\n+/g, '\n\n');
  
  // Split into sections: attributes and code
  const lines = cleanCode.split('\n');
  const attributeLines: string[] = [];
  const codeLines: string[] = [];
  
  let inAttributeSection = true;
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    
    if (inAttributeSection && trimmedLine.startsWith('Attribute VB_')) {
      attributeLines.push(line);
    } else if (trimmedLine.length === 0) {
      if (inAttributeSection) {
        attributeLines.push(line);
      } else {
        codeLines.push(line);
      }
    } else {
      inAttributeSection = false;
      codeLines.push(line);
    }
  }
  
  // Reconstruct the code with proper formatting
  let result = '';
  
  if (attributeLines.length > 0) {
    result += attributeLines.join('\n') + '\n';
    if (codeLines.length > 0) {
      result += '\n';
    }
  }
  
  if (codeLines.length > 0) {
    result += codeLines.join('\n');
  }
  
  return result.trim();
}

/**
 * Enhanced VBA code extraction from binary data
 */
export function extractVBACodeEnhanced(vbaData: Uint8Array, logger: LoggerCallback): VBAModule[] {
  try {
    logger('Starting enhanced VBA code extraction...', 'info');
    
    // Parse the VBA directory to find all modules
    const directoryEntries = parseVBADirectory(vbaData, logger);
    
    if (directoryEntries.length === 0) {
      logger('No VBA modules found in the project', 'warning');
      return [];
    }
    
    // Extract code for each module
    const modules: VBAModule[] = [];
    
    for (const entry of directoryEntries) {
      logger(`Extracting code for module: ${entry.name}`, 'info');
      
      const code = extractModuleCode(vbaData, entry.name, logger);
      const extractionSuccess = !code.startsWith('\'');
      
      modules.push({
        name: entry.name,
        type: entry.type,
        code,
        extractionSuccess
      });
      
      if (extractionSuccess) {
        logger(`Successfully extracted code for ${entry.name} (${code.split('\n').length} lines)`, 'success');
      } else {
        logger(`Failed to extract code for ${entry.name}`, 'warning');
      }
    }
    
    // Sort modules by type and name
    modules.sort((a, b) => {
      if (a.type !== b.type) {
        return a.type - b.type;
      }
      return a.name.localeCompare(b.name);
    });
    
    const successCount = modules.filter(m => m.extractionSuccess).length;
    logger(`Enhanced VBA extraction completed: ${successCount}/${modules.length} modules extracted successfully`, 
           successCount > 0 ? 'success' : 'warning');
    
    return modules;
    
  } catch (error) {
    logger(`Error in enhanced VBA code extraction: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return [];
  }
}

/**
 * Attempt to decompress VBA code if it's compressed
 */
export function decompressVBACode(compressedData: Uint8Array, logger: LoggerCallback): Uint8Array {
  try {
    // VBA code can be compressed using a simple RLE-like compression
    // This is a simplified decompression attempt
    
    const decompressed: number[] = [];
    let i = 0;
    
    while (i < compressedData.length) {
      const byte = compressedData[i];
      
      if (byte === 0x01) {
        // Potential compression marker
        if (i + 2 < compressedData.length) {
          const count = compressedData[i + 1];
          const value = compressedData[i + 2];
          
          for (let j = 0; j < count; j++) {
            decompressed.push(value);
          }
          
          i += 3;
        } else {
          decompressed.push(byte);
          i++;
        }
      } else {
        decompressed.push(byte);
        i++;
      }
    }
    
    const result = new Uint8Array(decompressed);
    logger(`VBA decompression: ${compressedData.length} -> ${result.length} bytes`, 'info');
    
    return result;
    
  } catch (error) {
    logger(`Error decompressing VBA code: ${error instanceof Error ? error.message : String(error)}`, 'warning');
    return compressedData; // Return original if decompression fails
  }
}