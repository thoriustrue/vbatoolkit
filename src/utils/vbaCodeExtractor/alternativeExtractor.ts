import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

/**
 * Alternative method to extract VBA modules when the primary method fails
 * @param workbook The workbook to extract VBA modules from
 * @param fileData The raw file data
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
export async function extractVBAModulesAlternative(
  workbook: XLSX.WorkBook, 
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  const modules: VBAModule[] = [];
  
  try {
    logger('Using alternative extraction method...', 'info');
    
    // Try to extract VBA modules directly from the ZIP structure
    try {
      const zip = await JSZip.loadAsync(fileData);
      const vbaProject = zip.file('xl/vbaProject.bin');
      
      if (vbaProject) {
        logger('Found vbaProject.bin in ZIP structure', 'info');
        const vbaContent = await vbaProject.async('uint8array');
        
        // Basic extraction of module names from vbaProject.bin
        const moduleNames = extractModuleNamesFromVBA(vbaContent, logger);
        
        if (moduleNames.length > 0) {
          logger(`Found ${moduleNames.length} module names in VBA project`, 'info');
          
          // Create placeholder modules with names but empty code
          for (const name of moduleNames) {
            modules.push({
              name,
              type: determineModuleType(name),
              code: `' Module: ${name}\n' Code could not be fully extracted\n' This is a placeholder for the module structure`
            });
          }
        } else {
          // If we can't extract module names, create a generic module
          modules.push({
            name: 'UnknownModule',
            type: VBAModuleType.Unknown,
            code: `' VBA code extraction was limited\n' The file contains VBA code but it could not be fully extracted`
          });
        }
      }
    } catch (zipError) {
      logger(`ZIP extraction failed: ${zipError instanceof Error ? zipError.message : String(zipError)}`, 'warning');
    }
    
    // If we have the raw VBA content from SheetJS, try to use it
    if (modules.length === 0 && workbook.vbaraw) {
      logger('Attempting to use raw VBA content from workbook', 'info');
      
      // Create a generic module with the information we have
      modules.push({
        name: 'ExtractedVBA',
        type: VBAModuleType.Standard,
        code: `' VBA code was detected but could not be fully extracted\n' The file contains VBA macros`
      });
    }
    
    if (modules.length === 0) {
      logger('Alternative extraction method could not extract any modules', 'warning');
    } else {
      logger(`Alternative extraction method extracted ${modules.length} module placeholders`, 'success');
    }
    
    return modules;
  } catch (error) {
    logger(`Error in alternative extraction method: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return modules;
  }
}

/**
 * Attempts to extract module names from the VBA binary content
 * @param vbaContent The VBA binary content
 * @param logger Callback function for logging messages
 * @returns An array of module names
 */
function extractModuleNamesFromVBA(vbaContent: Uint8Array, logger: LoggerCallback): string[] {
  const moduleNames: string[] = [];
  
  try {
    // Convert to string to search for module names
    const textDecoder = new TextDecoder('utf-8');
    const content = textDecoder.decode(vbaContent);
    
    // Look for common module name patterns
    // Module names are often stored with specific markers
    const modulePatterns = [
      /MODULE=([A-Za-z0-9_]+)/g,
      /Attribute VB_Name = "([^"]+)"/g,
      /\x00([A-Za-z0-9_]{2,30})\x00/g  // Module names between null bytes
    ];
    
    for (const pattern of modulePatterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        if (match[1] && !moduleNames.includes(match[1])) {
          // Filter out common false positives
          if (!match[1].includes('PROJECT') && 
              !match[1].includes('DOCUMENT') && 
              match[1].length > 1) {
            moduleNames.push(match[1]);
          }
        }
      }
    }
    
    // Look for common module names if we didn't find any
    if (moduleNames.length === 0) {
      const commonModuleNames = ['Module1', 'ThisWorkbook', 'Sheet1', 'UserForm1'];
      
      for (const name of commonModuleNames) {
        if (content.includes(name)) {
          moduleNames.push(name);
        }
      }
    }
    
    return moduleNames;
  } catch (error) {
    logger(`Error extracting module names: ${error instanceof Error ? error.message : String(error)}`, 'warning');
    return moduleNames;
  }
}

/**
 * Determines the module type based on its name
 * @param name The module name
 * @returns The module type
 */
function determineModuleType(name: string): VBAModuleType {
  if (name.startsWith('ThisWorkbook')) {
    return VBAModuleType.Document;
  } else if (name.startsWith('Sheet') || name.match(/^Worksheet\d+$/)) {
    return VBAModuleType.Document;
  } else if (name.startsWith('UserForm')) {
    return VBAModuleType.Form;
  } else if (name.startsWith('Class')) {
    return VBAModuleType.Class;
  } else {
    return VBAModuleType.Standard;
  }
} 