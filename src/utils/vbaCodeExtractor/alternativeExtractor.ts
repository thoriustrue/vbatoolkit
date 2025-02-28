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
        
        // Try to extract module information using multiple methods
        let moduleNames: string[] = [];
        
        // Method 1: Extract module names from vbaProject.bin
        moduleNames = extractModuleNamesFromVBA(vbaContent, logger);
        
        // Method 2: Try to extract from dir stream
        if (moduleNames.length === 0 || moduleNames.length < 5) { // If we found very few modules
          const dirStreamModules = extractModulesFromDirStream(vbaContent, logger);
          if (dirStreamModules.length > moduleNames.length) {
            moduleNames = dirStreamModules;
            logger(`Found ${moduleNames.length} modules using dir stream extraction`, 'info');
          }
        }
        
        // Method 3: Try to extract from module streams
        if (moduleNames.length === 0 || moduleNames.length < 5) {
          const streamModules = extractModulesFromStreams(vbaContent, logger);
          if (streamModules.length > moduleNames.length) {
            moduleNames = streamModules;
            logger(`Found ${moduleNames.length} modules using stream extraction`, 'info');
          }
        }
        
        // Method 4: Try to extract from sheet names
        if (moduleNames.length === 0 && workbook.SheetNames) {
          // Create modules for each sheet
          for (const sheetName of workbook.SheetNames) {
            moduleNames.push(`Sheet_${sheetName}`);
          }
          
          // Add ThisWorkbook
          moduleNames.push('ThisWorkbook');
          
          logger(`Created ${moduleNames.length} placeholder modules from sheet names`, 'info');
        }
        
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
      /\x00([A-Za-z0-9_]{2,30})\x00/g,  // Module names between null bytes
      /\x01([A-Za-z0-9_]{2,30})\x00/g,  // Another common pattern
      /([A-Za-z0-9_]{2,30})\.cls/gi,    // Class modules
      /([A-Za-z0-9_]{2,30})\.bas/gi,    // Standard modules
      /([A-Za-z0-9_]{2,30})\.frm/gi     // Form modules
    ];
    
    for (const pattern of modulePatterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        if (match[1] && !moduleNames.includes(match[1])) {
          // Filter out common false positives
          if (!match[1].includes('PROJECT') && 
              !match[1].includes('DOCUMENT') && 
              match[1].length > 1 &&
              !/^[0-9]+$/.test(match[1])) { // Exclude pure numbers
            moduleNames.push(match[1]);
          }
        }
      }
    }
    
    // Look for common module names if we didn't find any
    if (moduleNames.length === 0) {
      const commonModuleNames = [
        'Module1', 'Module2', 'Module3',
        'ThisWorkbook', 
        'Sheet1', 'Sheet2', 'Sheet3',
        'UserForm1', 'UserForm2',
        'Class1', 'Class2'
      ];
      
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
 * Attempts to extract module names from the dir stream in the VBA project
 * @param vbaContent The VBA binary content
 * @param logger Callback function for logging messages
 * @returns An array of module names
 */
function extractModulesFromDirStream(vbaContent: Uint8Array, logger: LoggerCallback): string[] {
  const moduleNames: string[] = [];
  
  try {
    // The dir stream contains information about all modules
    // Look for the dir stream marker
    const dirMarker = [0x44, 0x69, 0x72]; // "Dir" in ASCII
    
    for (let i = 0; i < vbaContent.length - dirMarker.length; i++) {
      let match = true;
      for (let j = 0; j < dirMarker.length; j++) {
        if (vbaContent[i + j] !== dirMarker[j]) {
          match = false;
          break;
        }
      }
      
      if (match) {
        // Found dir stream, now look for module records
        // Module records often start with specific patterns
        const searchRange = Math.min(5000, vbaContent.length - i);
        
        // Convert a section to string for easier searching
        const textDecoder = new TextDecoder('utf-8');
        const dirSection = textDecoder.decode(vbaContent.slice(i, i + searchRange));
        
        // Look for module name patterns in the dir stream
        const modulePattern = /([A-Za-z0-9_]{2,30})(\.bas|\.cls|\.frm)?[\x00-\x20]/g;
        let moduleMatch;
        
        while ((moduleMatch = modulePattern.exec(dirSection)) !== null) {
          const name = moduleMatch[1];
          if (name && 
              !moduleNames.includes(name) && 
              name.length > 1 && 
              !/^[0-9]+$/.test(name) &&
              !name.includes('PROJECT') &&
              !name.includes('VBA')) {
            moduleNames.push(name);
          }
        }
        
        // Only process the first dir stream we find
        break;
      }
    }
    
    return moduleNames;
  } catch (error) {
    logger(`Error extracting from dir stream: ${error instanceof Error ? error.message : String(error)}`, 'warning');
    return moduleNames;
  }
}

/**
 * Attempts to extract module names from individual module streams
 * @param vbaContent The VBA binary content
 * @param logger Callback function for logging messages
 * @returns An array of module names
 */
function extractModulesFromStreams(vbaContent: Uint8Array, logger: LoggerCallback): string[] {
  const moduleNames: string[] = [];
  
  try {
    // Module streams often contain "Attribute VB_Name" declarations
    const textDecoder = new TextDecoder('utf-8');
    const content = textDecoder.decode(vbaContent);
    
    // Split the content into potential module sections
    const sections = content.split(/[\x00-\x1F]{4,}/);
    
    for (const section of sections) {
      // Look for VB_Name attribute
      const nameMatch = section.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
      if (nameMatch && nameMatch[1]) {
        const name = nameMatch[1];
        if (!moduleNames.includes(name) && name.length > 1) {
          moduleNames.push(name);
        }
      }
      
      // Look for module headers
      const moduleMatch = section.match(/^(Module|Class|Form|Object)\s+([A-Za-z0-9_]+)/im);
      if (moduleMatch && moduleMatch[2]) {
        const name = moduleMatch[2];
        if (!moduleNames.includes(name) && name.length > 1) {
          moduleNames.push(name);
        }
      }
    }
    
    return moduleNames;
  } catch (error) {
    logger(`Error extracting from module streams: ${error instanceof Error ? error.message : String(error)}`, 'warning');
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
  } else if (name.startsWith('Sheet') || name.match(/^Worksheet\d+$/) || name.startsWith('Sheet_')) {
    return VBAModuleType.Document;
  } else if (name.startsWith('UserForm')) {
    return VBAModuleType.Form;
  } else if (name.startsWith('Class') || name.endsWith('.cls')) {
    return VBAModuleType.Class;
  } else {
    return VBAModuleType.Standard;
  }
} 