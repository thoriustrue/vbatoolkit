import { WorkBook } from 'xlsx';
import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';
import JSZip from 'jszip';

/**
 * Extracts VBA modules using alternative methods when primary extraction fails
 * @param workbook The workbook containing VBA
 * @param fileData The raw file data as ArrayBuffer
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
export async function extractVBAModulesAlternative(
  workbook: WorkBook,
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  const modules: VBAModule[] = [];
  
  try {
    logger('Attempting alternative VBA extraction method...', 'info');
    
    // Try to extract modules from the binary VBA content
    if (workbook.vbaraw) {
      logger('Found VBA binary content, attempting to extract module names...', 'info');
      
      // Extract module names from the VBA binary content
      const moduleNames = extractModuleNamesFromVBA(workbook.vbaraw);
      
      if (moduleNames.length > 0) {
        logger(`Found ${moduleNames.length} module names in VBA binary`, 'success');
        
        // Create placeholder modules for each name found
        for (const name of moduleNames) {
          const type = determineModuleType(name);
          modules.push({
            name,
            type,
            code: `' Module: ${name}\n' Code could not be fully extracted\n' This is a placeholder for the module structure`
          });
        }
      } else {
        logger('Could not find module names in VBA binary, trying ZIP extraction...', 'info');
      }
    }
    
    // If no modules found yet, try to extract from ZIP structure
    if (modules.length === 0) {
      try {
        const zip = await JSZip.loadAsync(fileData);
        
        // Look for vbaProject.bin
        const vbaProjectFile = zip.file('xl/vbaProject.bin');
        if (vbaProjectFile) {
          const vbaContent = await vbaProjectFile.async('uint8array');
          
          // Try to extract module names from the VBA project binary
          const moduleNames = extractModuleNamesFromVBA(vbaContent);
          
          if (moduleNames.length > 0) {
            logger(`Found ${moduleNames.length} module names in vbaProject.bin`, 'success');
            
            // Create placeholder modules for each name found
            for (const name of moduleNames) {
              const type = determineModuleType(name);
              modules.push({
                name,
                type,
                code: `' Module: ${name}\n' Code could not be fully extracted\n' This is a placeholder for the module structure`
              });
            }
          }
        }
        
        // If still no modules, check for sheet names in the workbook
        if (modules.length === 0 && workbook.SheetNames) {
          logger('Extracting sheet names as potential VBA modules...', 'info');
          
          // Add ThisWorkbook
          modules.push({
            name: 'ThisWorkbook',
            type: VBAModuleType.Document,
            code: `' Module: ThisWorkbook\n' Code could not be fully extracted\n' This is a placeholder for the module structure`
          });
          
          // Add sheets
          for (const sheetName of workbook.SheetNames) {
            modules.push({
              name: sheetName,
              type: VBAModuleType.Document,
              code: `' Module: ${sheetName}\n' Code could not be fully extracted\n' This is a placeholder for the module structure`
            });
          }
        }
      } catch (zipError) {
        logger(`Error extracting from ZIP: ${zipError instanceof Error ? zipError.message : String(zipError)}`, 'warning');
      }
    }
    
    // If still no modules found, create a generic module
    if (modules.length === 0) {
      logger('Could not extract any module information, creating generic module', 'warning');
      
      const UnknownModule: VBAModule = {
        name: 'ExtractedVBA',
        type: VBAModuleType.Unknown,
        code: '\'VBA code could not be extracted\n\'The file may have an unsupported format or corrupted VBA project'
      };
      
      modules.push(UnknownModule);
    }
    
    return modules;
  } catch (error) {
    logger(`Error in alternative extraction: ${error instanceof Error ? error.message : String(error)}`, 'error');
    
    // Return a generic module in case of error
    const ExtractedVBA: VBAModule = {
      name: 'ExtractedVBA',
      type: VBAModuleType.Standard,
      code: '\'Error during VBA extraction\n\'The file may have an unsupported format or corrupted VBA project'
    };
    
    return [ExtractedVBA];
  }
}

/**
 * Extracts module names from VBA binary content
 * @param vbaContent The VBA binary content
 * @returns An array of module names
 */
function extractModuleNamesFromVBA(vbaContent: Uint8Array): string[] {
  const moduleNames: string[] = [];
  
  try {
    // Convert binary to string for regex matching
    const vbaString = new TextDecoder('utf-8').decode(vbaContent);
    
    // Multiple regex patterns to find module names
    const patterns = [
      // Pattern for module headers
      /(?:Attribute\s+VB_Name\s*=\s*"([^"]+)")/gi,
      
      // Pattern for class modules
      /(?:BEGIN\s+(?:Class|Form|Module)\s+([A-Za-z0-9_]+))/gi,
      
      // Pattern for sheet modules
      /(?:Sheet([0-9]+))/gi,
      
      // Pattern for ThisWorkbook
      /(?:ThisWorkbook)/gi,
      
      // Pattern for UserForms
      /(?:UserForm([0-9]+))/gi,
      
      // Pattern for standard modules with common naming
      /(?:Module([0-9]+))/gi
    ];
    
    // Apply each pattern and collect unique names
    const foundNames = new Set<string>();
    
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(vbaString)) !== null) {
        if (match[1]) {
          foundNames.add(match[1]);
        } else if (match[0]) {
          foundNames.add(match[0]);
        }
      }
    }
    
    // Convert Set to Array
    moduleNames.push(...Array.from(foundNames));
  } catch (error) {
    // Silently fail and return empty array
    console.error('Error extracting module names:', error);
  }
  
  return moduleNames;
}

/**
 * Determines the type of a VBA module based on its name
 * @param moduleName The name of the module
 * @returns The type of the module
 */
function determineModuleType(moduleName: string): VBAModuleType {
  // Check for document modules
  if (moduleName === 'ThisWorkbook' || moduleName.startsWith('Sheet')) {
    return VBAModuleType.Document;
  }
  
  // Check for UserForms
  if (moduleName.includes('UserForm')) {
    return VBAModuleType.Form;
  }
  
  // Check for Class modules
  if (moduleName.includes('Class')) {
    return VBAModuleType.Class;
  }
  
  // Default to standard module
  return VBAModuleType.Standard;
} 