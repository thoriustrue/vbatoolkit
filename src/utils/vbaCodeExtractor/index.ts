import { readWorkbook } from '../xlsxWrapper';
import { LoggerCallback, ProgressCallback } from '../../types';
import { extractVBAModulesFromWorkbook, extractVBAModulesFromBinary, extractCodeFromModules } from './moduleExtractor';
import { extractVBAModulesAlternative } from './alternativeExtractor';
import { cleanAndDecodeVBACode } from './codeDecoder';
import { readFileAsArrayBuffer } from '../fileUtils';
import { VBAModule, VBAModuleType } from './types';
import * as XLSX from 'xlsx';

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
    
    // Try to extract VBA modules using multiple methods
    let modules: VBAModule[] = [];
    let extractionSuccess = false;
    
    // First attempt: Use SheetJS's built-in VBA extraction
    if (workbook.Workbook?.VBAProject) {
      logger('Extracting VBA modules using primary method...', 'info');
      modules = extractVBAModulesFromWorkbook(workbook, logger);
      
      if (modules.length > 0) {
        logger(`Successfully extracted ${modules.length} modules using primary method`, 'success');
        extractionSuccess = true;
      }
    }
    
    // If that didn't work, try alternative methods
    if (modules.length === 0) {
      logger('Primary extraction method failed. Trying alternative method...', 'info');
      modules = await extractVBAModulesAlternative(workbook, arrayBuffer, logger);
      
      if (modules.length > 0) {
        logger(`Successfully extracted ${modules.length} modules using alternative method`, 'info');
        // Mark as partial success since we only got module names
        extractionSuccess = true;
      }
    }
    
    if (modules.length === 0) {
      logger('No VBA modules could be extracted. The file may have an unsupported format or corrupted VBA project.', 'error');
      return { modules: [], success: false };
    }
    
    progressCallback(70);
    
    // Clean and decode the extracted modules
    modules = modules.map(module => ({
      ...module,
      code: cleanAndDecodeVBACode(module.code)
    }));
    
    // Sort modules by type and name for better organization
    modules.sort((a, b) => {
      // First sort by type priority
      const typePriority = {
        [VBAModuleType.Document]: 1,
        [VBAModuleType.Class]: 2,
        [VBAModuleType.Form]: 3,
        [VBAModuleType.Standard]: 4,
        [VBAModuleType.Unknown]: 5
      };
      
      const typeCompare = typePriority[a.type] - typePriority[b.type];
      if (typeCompare !== 0) return typeCompare;
      
      // Then sort by name
      return a.name.localeCompare(b.name);
    });
    
    logger(`VBA code extraction completed. Found ${modules.length} modules.`, 'success');
    progressCallback(100);
    
    return { modules, success: extractionSuccess };
  } catch (error) {
    logger(`Error during VBA code extraction: ${error instanceof Error ? error.message : String(error)}`, 'error');
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

/**
 * Creates a text file containing the process logs
 * @param logs Array of log messages
 * @param processType The type of process (e.g., "VBA Password Removal", "VBA Code Extraction")
 * @returns A Blob containing the logs
 */
export function createProcessLogsFile(logs: string[], processType: string): Blob {
  let content = `${processType} Process Logs\n`;
  content += `Date: ${new Date().toLocaleString()}\n`;
  content += `==========================================================\n\n`;
  
  logs.forEach(log => {
    content += `${log}\n`;
  });
  
  return new Blob([content], { type: 'text/plain' });
}

/**
 * Extracts VBA code modules from an Excel file
 * @param workbook The XLSX workbook
 * @param rawFileData The raw file data as Uint8Array
 * @param logger Callback function for logging messages
 * @returns Array of VBA modules
 */
export async function extractVBAModules(
  workbook: XLSX.WorkBook,
  rawFileData: Uint8Array,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  try {
    logger('Starting VBA code extraction process...', 'info');
    
    // Track extraction success
    let extractionSuccess = false;
    let modules: VBAModule[] = [];
    
    // Method 1: Try to extract modules from workbook
    logger('Attempting primary extraction method...', 'info');
    modules = extractVBAModulesFromWorkbook(workbook, logger);
    
    if (modules.length > 0 && modules.some(m => m.extractionSuccess)) {
      logger(`Successfully extracted ${modules.length} modules using primary method`, 'success');
      extractionSuccess = true;
    } else {
      logger('Primary extraction method did not yield complete results, trying alternative methods...', 'info');
      
      // Method 2: Try to extract modules from binary data
      if (workbook.vbaraw) {
        logger('Attempting extraction from VBA binary data...', 'info');
        const vbaData = new Uint8Array(workbook.vbaraw);
        
        // First try to extract module names
        const binaryModules = extractVBAModulesFromBinary(vbaData, logger);
        
        if (binaryModules.length > 0) {
          // Then try to extract code for each module
          const modulesWithCode = extractCodeFromModules(vbaData, binaryModules, logger);
          
          if (modulesWithCode.some(m => m.extractionSuccess)) {
            modules = modulesWithCode;
            logger(`Successfully extracted ${modulesWithCode.filter(m => m.extractionSuccess).length} modules using binary extraction`, 'success');
            extractionSuccess = true;
          } else {
            // Keep the module names at least
            modules = binaryModules;
            logger(`Extracted ${binaryModules.length} module names, but could not extract code`, 'warning');
          }
        }
      }
      
      // Method 3: Try alternative extraction method
      if (!extractionSuccess) {
        logger('Attempting alternative extraction method...', 'info');
        const alternativeModules = await extractVBAModulesAlternative(workbook, rawFileData, logger);
        
        if (alternativeModules.length > 0) {
          // If we already have module names but no code, merge the results
          if (modules.length > 0 && !modules.some(m => m.extractionSuccess)) {
            // Create a map of existing modules
            const moduleMap = new Map<string, VBAModule>();
            for (const module of modules) {
              moduleMap.set(module.name.toLowerCase(), module);
            }
            
            // Add any new modules from alternative extraction
            for (const altModule of alternativeModules) {
              const existingModule = moduleMap.get(altModule.name.toLowerCase());
              
              if (!existingModule) {
                modules.push(altModule);
              } else if (!existingModule.extractionSuccess && altModule.extractionSuccess) {
                // Update existing module with better code
                existingModule.code = altModule.code;
                existingModule.extractionSuccess = altModule.extractionSuccess;
              }
            }
            
            logger(`Combined results from multiple extraction methods, total modules: ${modules.length}`, 'info');
          } else {
            // Just use the alternative modules
            modules = alternativeModules;
            logger(`Extracted ${alternativeModules.length} modules using alternative method`, 'success');
          }
          
          extractionSuccess = modules.some(m => m.extractionSuccess);
        }
      }
    }
    
    // Sort modules by type and name for better organization
    modules.sort((a, b) => {
      // First sort by type
      if (a.type !== b.type) {
        return a.type - b.type;
      }
      // Then sort by name
      return a.name.localeCompare(b.name);
    });
    
    // Final status message
    if (modules.length === 0) {
      logger('No VBA modules could be extracted from this workbook', 'warning');
    } else if (!extractionSuccess) {
      logger(`Found ${modules.length} VBA modules, but could not extract code content`, 'warning');
    } else {
      const successCount = modules.filter(m => m.extractionSuccess).length;
      logger(`Successfully extracted ${successCount} out of ${modules.length} VBA modules`, 'success');
    }
    
    return modules;
  } catch (error) {
    logger(`Error during VBA code extraction: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return [];
  }
}

// Re-export types
export type { VBAModule, VBAModuleType } from './types'; 