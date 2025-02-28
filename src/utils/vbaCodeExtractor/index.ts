import { readWorkbook } from '../xlsxWrapper';
import { LoggerCallback, ProgressCallback } from '../../types';
import { extractVBAModulesFromWorkbook } from './moduleExtractor';
import { extractVBAModulesAlternative } from './alternativeExtractor';
import { cleanAndDecodeVBACode } from './codeDecoder';
import { readFileAsArrayBuffer } from '../fileUtils';
import { VBAModule, VBAModuleType } from './types';

/**
 * Extracts VBA code from an Excel file
 * @param file The Excel file to process
 * @param logger Callback function for logging messages
 * @param progressCallback Callback function for reporting progress
 * @returns A Promise that resolves to an object containing the extracted VBA code modules
 */
export async function extractVBACode(
  file: File,
  logger: LoggerCallback,
  progressCallback: ProgressCallback
): Promise<{ modules: VBAModule[], success: boolean }> {
  try {
    logger(`Processing file: ${file.name} for VBA code extraction`, 'info');
    progressCallback(10);
    
    // Read the file as an ArrayBuffer
    const arrayBuffer = await readFileAsArrayBuffer(file);
    
    logger('File loaded successfully. Analyzing Excel structure...', 'info');
    progressCallback(20);
    
    // Use SheetJS to read the workbook with VBA content
    const workbook = readWorkbook(arrayBuffer, { 
      bookVBA: true,  // Important: This tells SheetJS to preserve VBA
      cellFormula: false, // We don't need formulas
      cellHTML: false, // We don't need HTML
      cellText: false  // We don't need text conversion
    });
    
    // Check if the workbook has VBA
    if (!workbook.vbaraw) {
      logger('No VBA code found in this file. Make sure the file contains VBA macros.', 'error');
      return { modules: [], success: false };
    }
    
    logger('VBA project found in the workbook.', 'success');
    progressCallback(40);
    
    // Try to extract VBA modules using SheetJS
    let modules: VBAModule[] = [];
    
    // First attempt: Use SheetJS's built-in VBA extraction
    if (workbook.Workbook?.VBAProject) {
      logger('Extracting VBA modules using primary method...', 'info');
      modules = extractVBAModulesFromWorkbook(workbook, logger);
    }
    
    // If that didn't work, try alternative methods
    if (modules.length === 0) {
      logger('Primary extraction method failed. Trying alternative method...', 'info');
      modules = await extractVBAModulesAlternative(workbook, arrayBuffer, logger);
    }
    
    if (modules.length === 0) {
      logger('No VBA modules could be extracted. The file may have an unsupported format or corrupted VBA project.', 'error');
      return { modules: [], success: false };
    }
    
    // Clean and decode the extracted modules
    modules = modules.map(module => ({
      ...module,
      code: cleanAndDecodeVBACode(module.code)
    }));
    
    logger(`Successfully extracted ${modules.length} VBA module(s).`, 'success');
    progressCallback(100);
    
    return { modules, success: true };
  } catch (error) {
    logger(`Error extracting VBA code: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return { modules: [], success: false };
  }
}

/**
 * Creates a text file containing all the VBA code from the extracted modules
 * @param modules The VBA modules to include in the file
 * @param fileName The name of the original Excel file
 * @returns A Blob containing the VBA code
 */
export function createVBACodeFile(modules: VBAModule[], fileName: string): Blob {
  let content = `VBA Code extracted from: ${fileName}\n`;
  content += `Extraction date: ${new Date().toLocaleString()}\n`;
  content += `Number of modules: ${modules.length}\n\n`;
  
  modules.forEach(module => {
    content += `'==========================================================\n`;
    content += `' Module: ${module.name}\n`;
    content += `' Type: ${module.type}\n`;
    content += `'==========================================================\n\n`;
    content += module.code;
    content += '\n\n';
  });
  
  return new Blob([content], { type: 'text/plain' });
}

// Re-export types
export type { VBAModule, VBAModuleType } from './types'; 