import { readWorkbook } from './xlsxWrapper';
import { utils as XLSXUtils } from 'xlsx';
import JSZip from 'jszip';
import { LoggerCallback } from './types';

// Type for the logger callback function
type ProgressCallback = (progress: number) => void;

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
    
    // Load ZIP file
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // Check if file contains VBA project
    const vbaProject = zip.file('xl/vbaProject.bin');
    if (!vbaProject) {
      logger('No VBA project found in the file', 'warning');
      return { modules: [], success: false };
    }
    
    logger('VBA project found in the file.', 'success');
    progressCallback(40);
    
    // Extract VBA project binary
    const vbaContent = await vbaProject.async('uint8array');
    logger(`VBA project size: ${vbaContent.length} bytes`, 'info');
    
    // Extract module information from dir stream
    const modules = await extractModulesFromVBAProject(vbaContent, logger);
    
    if (modules.length === 0) {
      logger('No VBA modules found in the project', 'warning');
      return { modules: [], success: false };
    }
    
    logger(`Successfully extracted ${modules.length} VBA modules`, 'success');
    progressCallback(100);
    
    return { modules, success: true };
  } catch (error) {
    logger(`Error extracting VBA code: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return { modules: [], success: false };
  }
}

/**
 * Represents a VBA code module
 */
export interface VBAModule {
  name: string;
  code: string;
  type: 'standard' | 'class' | 'form' | 'document' | 'unknown';
}

/**
 * Extracts modules from VBA project binary
 */
async function extractModulesFromVBAProject(
  vbaContent: Uint8Array,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  const modules: VBAModule[] = [];
  
  try {
    // Find the dir stream in the VBA project
    const dirSignature = new Uint8Array([0x44, 0x69, 0x72]); // "Dir" in ASCII
    let dirOffset = findSignature(vbaContent, dirSignature);
    
    if (dirOffset === -1) {
      logger('Could not find dir stream in VBA project', 'error');
      return modules;
    }
    
    // Skip dir header (typically 4 bytes)
    dirOffset += dirSignature.length + 4;
    
    // Parse the dir stream to find module information
    const moduleInfos = parseModuleInfos(vbaContent, dirOffset, logger);
    
    // Extract each module's code
    for (const moduleInfo of moduleInfos) {
      // Find the module stream
      const moduleNameBytes = new TextEncoder().encode(moduleInfo.name);
      let moduleOffset = findSignature(vbaContent, moduleNameBytes);
      
      if (moduleOffset === -1) {
        logger(`Could not find module stream for ${moduleInfo.name}`, 'warning');
        continue;
      }
      
      // Skip to the actual code (typically starts after module name and some headers)
      moduleOffset += moduleNameBytes.length + moduleInfo.headerSize;
      
      // Extract the code with proper encoding handling
      const codeBytes = vbaContent.slice(moduleOffset, moduleOffset + moduleInfo.codeSize);
      let code = '';
      
      // Try multiple encodings to handle international characters correctly
      try {
        // First try UTF-8
        code = new TextDecoder('utf-8').decode(codeBytes);
      } catch (e) {
        try {
          // Fall back to Windows-1252 (common for VBA)
          code = new TextDecoder('windows-1252').decode(codeBytes);
        } catch (e2) {
          // Last resort: ISO-8859-1 (Latin-1)
          code = new TextDecoder('iso-8859-1').decode(codeBytes);
        }
      }
      
      // Clean up the code
      code = cleanVBACode(code);
      
      modules.push({
        name: moduleInfo.name,
        code: code,
        type: determineModuleType(moduleInfo.name, code)
      });
      
      logger(`Extracted module: ${moduleInfo.name} (${moduleInfo.codeSize} bytes)`, 'info');
    }
    
    return modules;
  } catch (error) {
    logger(`Error parsing VBA project: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return modules;
  }
}

/**
 * Parses module information from the dir stream
 */
function parseModuleInfos(
  vbaContent: Uint8Array,
  dirOffset: number,
  logger: LoggerCallback
): Array<{name: string, codeSize: number, headerSize: number}> {
  const moduleInfos: Array<{name: string, codeSize: number, headerSize: number}> = [];
  
  try {
    // The dir stream contains records for each module
    // Each record has a fixed structure with name, type, and offset information
    
    let offset = dirOffset;
    
    // Process records until we reach the end of the dir stream or find too many modules
    const maxModules = 100; // Safety limit
    for (let i = 0; i < maxModules; i++) {
      // Check if we've reached the end of the dir stream
      if (offset + 10 >= vbaContent.length) break;
      
      // Read record ID (2 bytes)
      const recordId = vbaContent[offset] | (vbaContent[offset + 1] << 8);
      offset += 2;
      
      // Module records have ID 0x0031
      if (recordId !== 0x0031) {
        // Skip non-module records
        const size = vbaContent[offset] | (vbaContent[offset + 1] << 8);
        offset += 2 + size;
        continue;
      }
      
      // Read record size (2 bytes)
      const recordSize = vbaContent[offset] | (vbaContent[offset + 1] << 8);
      offset += 2;
      
      // Skip to the name length (typically at offset +6)
      offset += 6;
      
      // Read name length (2 bytes)
      const nameLength = vbaContent[offset] | (vbaContent[offset + 1] << 8);
      offset += 2;
      
      // Read module name
      let name = '';
      for (let j = 0; j < nameLength && offset + j < vbaContent.length; j++) {
        name += String.fromCharCode(vbaContent[offset + j]);
      }
      offset += nameLength;
      
      // Skip to code size (typically at offset +8 from name end)
      offset += 8;
      
      // Read code size (4 bytes)
      const codeSize = vbaContent[offset] | 
                      (vbaContent[offset + 1] << 8) | 
                      (vbaContent[offset + 2] << 16) | 
                      (vbaContent[offset + 3] << 24);
      offset += 4;
      
      // Estimate header size (typically around 24 bytes)
      const headerSize = 24;
      
      moduleInfos.push({
        name,
        codeSize,
        headerSize
      });
      
      // Skip to the next record
      offset += recordSize - 22 - nameLength;
    }
    
    return moduleInfos;
  } catch (error) {
    logger(`Error parsing module information: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return moduleInfos;
  }
}

/**
 * Finds a signature in a byte array
 */
function findSignature(data: Uint8Array, signature: Uint8Array): number {
  for (let i = 0; i <= data.length - signature.length; i++) {
    let found = true;
    for (let j = 0; j < signature.length; j++) {
      if (data[i + j] !== signature[j]) {
        found = false;
        break;
      }
    }
    if (found) return i;
  }
  return -1;
}

/**
 * Cleans up extracted VBA code
 */
function cleanVBACode(code: string): string {
  // Remove null characters
  code = code.replace(/\0/g, '');
  
  // Fix line endings
  code = code.replace(/\r\n|\r|\n/g, '\r\n');
  
  // Remove any binary garbage at the beginning or end
  code = code.replace(/^[^\w\s'"(]+/, '');
  code = code.replace(/[^\w\s'")\r\n;]+$/, '');
  
  // Fix common encoding issues
  code = code.replace(/â€œ/g, '"');  // Smart quotes
  code = code.replace(/â€/g, '"');   // Smart quotes
  code = code.replace(/â€™/g, "'");  // Smart apostrophe
  code = code.replace(/â€"/g, "-");  // Em dash
  
  return code;
}

/**
 * Determines the type of VBA module
 */
function determineModuleType(name: string, code: string): VBAModule['type'] {
  // Check for class modules
  if (code.includes('Option Explicit\r\n\r\nAttribute VB_Name =') || 
      name.endsWith('.cls') || 
      code.includes('Attribute VB_Exposed = ')) {
    return 'class';
  }
  
  // Check for form modules
  if (name.endsWith('.frm') || 
      code.includes('Begin VB.Form') || 
      code.includes('Attribute VB_Exposed = True')) {
    return 'form';
  }
  
  // Check for document modules
  if (name === 'ThisWorkbook' || 
      name.startsWith('Sheet') || 
      code.includes('Attribute VB_VarHelpID = ')) {
    return 'document';
  }
  
  // Default to standard module
  return 'standard';
}

/**
 * Reads a File as an ArrayBuffer
 * @param file The file to read
 * @returns A Promise that resolves to an ArrayBuffer
 */
function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      if (e.target?.result instanceof ArrayBuffer) {
        resolve(e.target.result);
      } else {
        reject(new Error('Failed to read file as ArrayBuffer'));
      }
    };
    reader.onerror = () => reject(new Error('File read error'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Creates a downloadable file with the extracted VBA code
 * @param modules The VBA modules to include in the file
 * @param originalFileName The name of the original Excel file
 * @returns A Blob containing the VBA code
 */
export function createVBACodeFile(modules: VBAModule[], originalFileName: string): Blob {
  // Create a text file with all the VBA code
  let content = `VBA Code extracted from: ${originalFileName}\n`;
  content += `Extraction date: ${new Date().toLocaleString()}\n\n`;
  content += `Total modules: ${modules.length}\n\n`;
  
  for (const module of modules) {
    content += `'==========================================================\n`;
    content += `' Module: ${module.name}\n`;
    content += `' Type: ${module.type}\n`;
    content += `'==========================================================\n\n`;
    content += `${module.code}\n\n\n`;
  }
  
  return new Blob([content], { type: 'text/plain' });
}