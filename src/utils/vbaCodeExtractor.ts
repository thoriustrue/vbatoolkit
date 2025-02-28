import { readWorkbook } from './xlsxWrapper';
import { utils as XLSXUtils } from 'xlsx';

// Type for the logger callback function
type LoggerCallback = (message: string, type: 'info' | 'error' | 'success') => void;
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
 * Represents a VBA code module
 */
export interface VBAModule {
  name: string;
  type: VBAModuleType;
  code: string;
}

/**
 * Types of VBA modules
 */
export enum VBAModuleType {
  Standard = 'Standard Module',
  Class = 'Class Module',
  Form = 'UserForm',
  Document = 'Document Module',
  Unknown = 'Unknown'
}

/**
 * Extracts VBA modules from a workbook using SheetJS
 * @param workbook The workbook to extract VBA modules from
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
function extractVBAModulesFromWorkbook(workbook: XLSX.WorkBook, logger: LoggerCallback): VBAModule[] {
  const modules: VBAModule[] = [];
  
  try {
    // Access the VBA project
    if (!workbook.Workbook?.VBAProject) {
      return modules;
    }
    
    const vbaProject = workbook.Workbook.VBAProject;
    
    // Get the module names from the VBA project
    const moduleNames = Object.keys(vbaProject);
    
    for (const moduleName of moduleNames) {
      // Skip non-module properties
      if (moduleName === 'Name' || moduleName === 'CodeName' || typeof vbaProject[moduleName] !== 'string') {
        continue;
      }
      
      // Get the module code
      const moduleCode = vbaProject[moduleName].toString();
      
      if (moduleCode.trim()) {
        // Determine module type based on content
        let moduleType = VBAModuleType.Unknown;
        
        if (moduleCode.includes('Attribute VB_Name = "')) {
          if (moduleCode.includes('Attribute VB_Base = "0{')) {
            moduleType = VBAModuleType.Form;
          } else if (moduleCode.includes('Attribute VB_PredeclaredId = True')) {
            if (moduleCode.includes('Attribute VB_Exposed = True')) {
              moduleType = VBAModuleType.Document;
            } else {
              moduleType = VBAModuleType.Standard;
            }
          } else if (moduleCode.includes('Attribute VB_GlobalNameSpace = False') && 
                    moduleCode.includes('Attribute VB_Creatable = False')) {
            moduleType = VBAModuleType.Class;
          } else {
            moduleType = VBAModuleType.Standard;
          }
        }
        
        modules.push({
          name: moduleName,
          type: moduleType,
          code: moduleCode
        });
        
        logger(`Extracted module: ${moduleName} (${moduleType})`, 'info');
      }
    }
    
    return modules;
  } catch (error) {
    logger(`Error in primary extraction method: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return modules;
  }
}

/**
 * Alternative method to extract VBA modules when the primary method fails
 * @param workbook The workbook to extract VBA modules from
 * @param fileData The raw file data
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
async function extractVBAModulesAlternative(
  workbook: XLSX.WorkBook, 
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  const modules: VBAModule[] = [];
  
  try {
<<<<<<< HEAD
<<<<<<< HEAD
    // Find the dir stream in the VBA project - try multiple signatures
    const dirSignatures = [
      new Uint8Array([0x44, 0x69, 0x72]), // "Dir" in ASCII
      new Uint8Array([0x01, 0x44, 0x69, 0x72]), // Prefixed "Dir"
      new Uint8Array([0x44, 0x49, 0x52]) // "DIR" in uppercase
    ];
    
    let dirOffset = -1;
    for (const signature of dirSignatures) {
      dirOffset = findSignature(vbaContent, signature);
      if (dirOffset !== -1) {
        logger(`Found dir stream at offset ${dirOffset} with signature ${Array.from(signature).map(b => b.toString(16).padStart(2, '0')).join(' ')}`, 'info');
        dirOffset += signature.length;
        break;
      }
    }
=======
    // Find the dir stream in the VBA project
    const dirSignature = new Uint8Array([0x44, 0x69, 0x72]); // "Dir" in ASCII
    let dirOffset = findSignature(vbaContent, dirSignature);
>>>>>>> parent of 998746b (Fixes)
    
    if (dirOffset === -1) {
      logger('Could not find dir stream in VBA project', 'error');
      return modules;
    }
    
    // Skip dir header (typically 4 bytes)
<<<<<<< HEAD
    dirOffset += 4;
=======
    // If the workbook has vbaraw but we couldn't extract modules using the primary method,
    // try to parse the vbaraw data directly
    if (!workbook.vbaraw) {
      return modules;
    }
=======
    dirOffset += dirSignature.length + 4;
>>>>>>> parent of 998746b (Fixes)
    
    logger('Analyzing VBA project structure...', 'info');
>>>>>>> parent of ef6e378 (Update vbaCodeExtractor.ts)
    
    // Convert vbaraw to Uint8Array for analysis
    const vbaData = new Uint8Array(workbook.vbaraw);
    
    // Look for module name patterns in the VBA data
    // Module names are typically preceded by "Attribute VB_Name = " in ASCII
    const namePattern = [0x41, 0x74, 0x74, 0x72, 0x69, 0x62, 0x75, 0x74, 0x65, 0x20, 0x56, 0x42, 0x5F, 0x4E, 0x61, 0x6D, 0x65, 0x20, 0x3D, 0x20, 0x22]; // "Attribute VB_Name = ""
    
    const nameIndices = findAllPatterns(vbaData, namePattern);
    logger(`Found ${nameIndices.length} potential VBA module(s).`, 'info');
    
    for (let i = 0; i < nameIndices.length; i++) {
      const nameStart = nameIndices[i] + namePattern.length;
      let nameEnd = nameStart;
      
      // Find the closing quote of the module name
      while (nameEnd < vbaData.length && vbaData[nameEnd] !== 0x22) {
        nameEnd++;
      }
      
      if (nameEnd < vbaData.length) {
        // Extract the module name
        const nameBytes = vbaData.slice(nameStart, nameEnd);
        const moduleName = new TextDecoder().decode(nameBytes);
        
        // Find the end of this module (start of next module or end of file)
        const nextModuleIndex = i + 1 < nameIndices.length ? nameIndices[i + 1] : vbaData.length;
        
        // Extract the module code
        // We need to find where the actual code starts after the attributes
        let codeStart = nameIndices[i];
        let attributesEnd = -1;
        
        // Look for the end of attributes section (typically ends with a line break after the last attribute)
        for (let j = nameIndices[i]; j < nextModuleIndex - 20; j++) {
          // Check for "Attribute VB_" pattern
          if (vbaData[j] === 0x41 && vbaData[j+1] === 0x74 && vbaData[j+2] === 0x74 && 
              vbaData[j+3] === 0x72 && vbaData[j+4] === 0x69 && vbaData[j+5] === 0x62 && 
              vbaData[j+6] === 0x75 && vbaData[j+7] === 0x74 && vbaData[j+8] === 0x65 && 
              vbaData[j+9] === 0x20 && vbaData[j+10] === 0x56 && vbaData[j+11] === 0x42 && 
              vbaData[j+12] === 0x5F) {
            // Found an attribute, update the potential end of attributes section
            // Look for the end of this line (CR or LF)
            let lineEnd = j;
            while (lineEnd < nextModuleIndex && vbaData[lineEnd] !== 0x0D && vbaData[lineEnd] !== 0x0A) {
              lineEnd++;
            }
            
            // Skip any CR/LF characters
            while (lineEnd < nextModuleIndex && (vbaData[lineEnd] === 0x0D || vbaData[lineEnd] === 0x0A)) {
              lineEnd++;
            }
            
            attributesEnd = lineEnd;
          }
        }
        
        // If we found the end of attributes, set the code start to that position
        if (attributesEnd !== -1) {
          codeStart = attributesEnd;
        }
        
        // Extract the module code
        const moduleBytes = vbaData.slice(codeStart, nextModuleIndex);
        let moduleCode = new TextDecoder().decode(moduleBytes);
        
        // Clean up the code - remove null bytes and control characters
        moduleCode = moduleCode.replace(/\0/g, '').replace(/[\x01-\x09\x0B\x0C\x0E-\x1F]/g, '');
        
        // Determine module type
        let moduleType = VBAModuleType.Unknown;
        
        if (moduleCode.includes('Attribute VB_Base = "0{')) {
          moduleType = VBAModuleType.Form;
        } else if (moduleCode.includes('Attribute VB_PredeclaredId = True')) {
          if (moduleCode.includes('Attribute VB_Exposed = True')) {
            moduleType = VBAModuleType.Document;
          } else {
            moduleType = VBAModuleType.Standard;
          }
        } else if (moduleCode.includes('Attribute VB_GlobalNameSpace = False') && 
                  moduleCode.includes('Attribute VB_Creatable = False')) {
          moduleType = VBAModuleType.Class;
        } else {
          moduleType = VBAModuleType.Standard;
        }
        
        modules.push({
          name: moduleName,
          type: moduleType,
          code: moduleCode
        });
        
        logger(`Extracted module: ${moduleName} (${moduleType})`, 'info');
      }
    }
    
    // If we still couldn't extract any modules, try one more approach
    if (modules.length === 0) {
      logger('Trying deep extraction method...', 'info');
      
      // Try to find code blocks directly
      const codePatterns = [
        [0x53, 0x75, 0x62, 0x20], // "Sub "
        [0x46, 0x75, 0x6E, 0x63, 0x74, 0x69, 0x6F, 0x6E, 0x20], // "Function "
        [0x50, 0x75, 0x62, 0x6C, 0x69, 0x63, 0x20, 0x53, 0x75, 0x62, 0x20], // "Public Sub "
        [0x50, 0x75, 0x62, 0x6C, 0x69, 0x63, 0x20, 0x46, 0x75, 0x6E, 0x63, 0x74, 0x69, 0x6F, 0x6E, 0x20], // "Public Function "
        [0x50, 0x72, 0x69, 0x76, 0x61, 0x74, 0x65, 0x20, 0x53, 0x75, 0x62, 0x20], // "Private Sub "
        [0x50, 0x72, 0x69, 0x76, 0x61, 0x74, 0x65, 0x20, 0x46, 0x75, 0x6E, 0x63, 0x74, 0x69, 0x6F, 0x6E, 0x20] // "Private Function "
      ];
      
      let codeBlocks: string[] = [];
      
      for (const pattern of codePatterns) {
        const patternIndices = findAllPatterns(vbaData, pattern);
        
        for (const index of patternIndices) {
          // Extract a reasonable chunk of code (up to 10KB)
          const codeChunkSize = 10240;
          const codeChunk = vbaData.slice(index, Math.min(index + codeChunkSize, vbaData.length));
          
          // Convert to string and clean up
          let codeBlock = new TextDecoder().decode(codeChunk);
          codeBlock = codeBlock.replace(/\0/g, '').replace(/[\x01-\x09\x0B\x0C\x0E-\x1F]/g, '');
          
          // Try to find the end of the procedure
          const endIndex = codeBlock.indexOf('End Sub');
          if (endIndex !== -1) {
            codeBlock = codeBlock.substring(0, endIndex + 7);
          } else {
            const endFunctionIndex = codeBlock.indexOf('End Function');
            if (endFunctionIndex !== -1) {
              codeBlock = codeBlock.substring(0, endFunctionIndex + 12);
            }
          }
          
          codeBlocks.push(codeBlock);
        }
      }
      
      // If we found any code blocks, create a single module with all the code
      if (codeBlocks.length > 0) {
        // Remove duplicates
        codeBlocks = [...new Set(codeBlocks)];
        
        modules.push({
          name: 'ExtractedCode',
          type: VBAModuleType.Unknown,
          code: codeBlocks.join('\n\n')
        });
        
        logger(`Extracted ${codeBlocks.length} code block(s) into a single module.`, 'info');
      }
    }
    
    return modules;
  } catch (error) {
    logger(`Error in alternative extraction method: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return modules;
  }
}

/**
 * Cleans and decodes VBA code to make it more readable
 * @param code The raw VBA code to clean and decode
 * @returns The cleaned and decoded VBA code
 */
function cleanAndDecodeVBACode(code: string): string {
  // Remove null bytes and control characters
  let cleanedCode = code.replace(/\0/g, '').replace(/[\x01-\x08\x0B\x0C\x0E-\x1F]/g, '');
  
  // Fix common encoding issues
  cleanedCode = cleanedCode
    // Replace common Unicode replacement characters
    .replace(/\uFFFD/g, '')
    // Fix line endings
    .replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    // Remove excessive line breaks
    .replace(/\n{3,}/g, '\n\n')
    // Fix common VBA attribute lines
    .replace(/Attribute VB_\w+ = [^\n]+\n?/g, (match) => {
      // Keep attribute lines but ensure they end with a newline
      return match.endsWith('\n') ? match : match + '\n';
    });
  
  // Extract the actual code part (after attributes)
  const attributeEndIndex = cleanedCode.lastIndexOf('Attribute VB_');
  if (attributeEndIndex !== -1) {
    const lineEndIndex = cleanedCode.indexOf('\n', attributeEndIndex);
    if (lineEndIndex !== -1) {
      const codeStartIndex = lineEndIndex + 1;
      // Keep attributes but ensure there's a blank line before the code
      cleanedCode = cleanedCode.substring(0, codeStartIndex) + '\n' + cleanedCode.substring(codeStartIndex);
    }
  }
  
  // Fix common VBA syntax issues
  cleanedCode = cleanedCode
    // Fix broken string concatenation
    .replace(/" \+/g, '" &')
    .replace(/\+ "/g, '& "')
    // Fix broken line continuation
    .replace(/_\s+\n/g, ' _\n')
    // Fix common VBA keywords that might be corrupted
    .replace(/Dim(\w)/g, 'Dim $1')
    .replace(/Set(\w)/g, 'Set $1')
    .replace(/If(\w)/g, 'If $1')
    .replace(/End(\w)/g, (match, p1) => {
      // Only add space if the next word is a known VBA block end
      if (['Sub', 'Function', 'If', 'With', 'Select', 'Type', 'Property'].includes(p1)) {
        return 'End ' + p1;
      }
      return match;
    });
  
  // Remove any remaining non-printable characters
  cleanedCode = cleanedCode.replace(/[^\x20-\x7E\n]/g, '');
  
  return cleanedCode;
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
 * Finds all occurrences of a byte pattern in a Uint8Array
 * @param data The data to search in
 * @param pattern The pattern to search for
 * @returns An array of indices where the pattern was found
 */
function findAllPatterns(data: Uint8Array, pattern: number[]): number[] {
  const indices: number[] = [];
  
  for (let i = 0; i <= data.length - pattern.length; i++) {
    let found = true;
    for (let j = 0; j < pattern.length; j++) {
      if (data[i + j] !== pattern[j]) {
        found = false;
        break;
      }
    }
    if (found) {
      indices.push(i);
    }
  }
  
  return indices;
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