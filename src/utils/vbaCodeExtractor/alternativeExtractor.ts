import * as XLSX from 'xlsx';
import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';
import JSZip from 'jszip';

/**
 * Alternative method to extract VBA modules from a workbook
 * This method uses direct ZIP extraction to access the vbaProject.bin file
 * @param workbook The XLSX workbook
 * @param rawFileData The raw file data as Uint8Array
 * @param logger Callback function for logging messages
 * @returns Array of VBA modules
 */
export async function extractVBAModulesAlternative(
  workbook: XLSX.WorkBook,
  rawFileData: Uint8Array,
  logger: LoggerCallback
): Promise<VBAModule[]> {
  try {
    logger('Starting alternative VBA module extraction...', 'info');
    
    const modules: VBAModule[] = [];
    
    // Check if we have binary VBA content
    if (!workbook.vbaraw) {
      logger('No VBA binary content found in workbook', 'warning');
      return modules;
    }
    
    // Try to extract module names from the binary content
    logger('Attempting to extract module names from binary content...', 'info');
    
    const vbaContent = new TextDecoder('utf-8').decode(new Uint8Array(workbook.vbaraw));
    
    // Method 1: Extract module names using VB_Name attribute
    const moduleNameMatches = vbaContent.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/g);
    if (moduleNameMatches && moduleNameMatches.length > 0) {
      logger(`Found ${moduleNameMatches.length} module names using VB_Name attribute`, 'info');
      
      for (const match of moduleNameMatches) {
        const nameMatch = match.match(/"([^"]+)"/);
        if (nameMatch && nameMatch[1]) {
          const name = nameMatch[1];
          
          // Determine module type based on name and surrounding content
          let type = VBAModuleType.Standard;
          const moduleStart = vbaContent.indexOf(match);
          const moduleContext = vbaContent.substring(
            Math.max(0, moduleStart - 100),
            Math.min(vbaContent.length, moduleStart + 500)
          );
          
          if (name.toLowerCase() === 'thisdocument' || name.toLowerCase() === 'thisworkbook') {
            type = VBAModuleType.Document;
          } else if (name.toLowerCase().startsWith('sheet') || name.toLowerCase().includes('worksheet')) {
            type = VBAModuleType.Document;
          } else if (moduleContext.includes('Attribute VB_Creatable = False') && 
                    moduleContext.includes('Attribute VB_GlobalNameSpace = False')) {
            type = VBAModuleType.Class;
          } else if (moduleContext.includes('Begin VB.Form') || name.toLowerCase().includes('form')) {
            type = VBAModuleType.Form;
          }
          
          // Try to extract code for this module
          let code = '';
          let extractionSuccess = false;
          
          // Find the module's code section
          const codeStartIndex = vbaContent.indexOf(match);
          if (codeStartIndex >= 0) {
            // Find the next module or end of content
            let nextModuleIndex = vbaContent.length;
            for (const otherMatch of moduleNameMatches) {
              if (otherMatch !== match) {
                const otherIndex = vbaContent.indexOf(otherMatch);
                if (otherIndex > codeStartIndex && otherIndex < nextModuleIndex) {
                  nextModuleIndex = otherIndex;
                }
              }
            }
            
            // Extract the code between this module and the next
            let moduleCode = vbaContent.substring(codeStartIndex, nextModuleIndex).trim();
            
            // Clean up the code - remove attributes and binary artifacts
            const attributeEndIndex = moduleCode.search(/(?:^|\r\n)(?!Attribute VB_)/m);
            if (attributeEndIndex > 0) {
              code = moduleCode.substring(attributeEndIndex).trim();
              code = code.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
              extractionSuccess = code.length > 0;
            }
          }
          
          if (!extractionSuccess) {
            code = `' Code could not be fully extracted for module: ${name}`;
          }
          
          modules.push({
            name,
            type,
            code,
            extractionSuccess
          });
          
          logger(`Extracted module: ${name} (${VBAModuleType[type]})`, extractionSuccess ? 'success' : 'warning');
        }
      }
    }
    
    // Method 2: If no modules found, try ZIP extraction
    if (modules.length === 0) {
      logger('No modules found using binary content, trying ZIP extraction...', 'info');
      
      try {
        // Load the file as a ZIP archive
        const zip = await JSZip.loadAsync(rawFileData);
        
        // Look for vbaProject.bin
        const vbaProjectFile = zip.file(/xl\/vbaProject.bin$/i)[0];
        
        if (vbaProjectFile) {
          logger('Found vbaProject.bin in ZIP structure', 'info');
          
          // Extract the file content
          const vbaProjectContent = await vbaProjectFile.async('uint8array');
          const vbaText = new TextDecoder('utf-8').decode(vbaProjectContent);
          
          // Look for module names
          const dirMatches = vbaText.match(/(?:MODULE|CLASS|DOCUMENT)=([^\r\n]+)/g) || [];
          
          for (const match of dirMatches) {
            let type = VBAModuleType.Unknown;
            let name = '';
            
            if (match.startsWith('MODULE=')) {
              type = VBAModuleType.Standard;
              name = match.substring(7).trim();
            } else if (match.startsWith('CLASS=')) {
              type = VBAModuleType.Class;
              name = match.substring(6).trim();
            } else if (match.startsWith('DOCUMENT=')) {
              type = VBAModuleType.Document;
              name = match.substring(9).trim();
            }
            
            if (name) {
              modules.push({
                name,
                type,
                code: `' Code could not be fully extracted for module: ${name}`,
                extractionSuccess: false
              });
              
              logger(`Found module via ZIP extraction: ${name} (${VBAModuleType[type]})`, 'info');
            }
          }
          
          // Look for form modules
          const formMatches = vbaText.match(/Begin VB\.Form ([^\r\n]+)/g) || [];
          for (const match of formMatches) {
            const name = match.substring(14).trim();
            if (name) {
              modules.push({
                name,
                type: VBAModuleType.Form,
                code: `' Code could not be fully extracted for form: ${name}`,
                extractionSuccess: false
              });
              
              logger(`Found form module via ZIP extraction: ${name}`, 'info');
            }
          }
        }
      } catch (zipError) {
        logger(`Error during ZIP extraction: ${zipError instanceof Error ? zipError.message : String(zipError)}`, 'warning');
      }
    }
    
    // Method 3: Try to find module names in the workbook structure
    if (modules.length === 0 && workbook.Workbook) {
      logger('Attempting to find modules in workbook structure...', 'info');
      
      // Check for ThisWorkbook
      modules.push({
        name: 'ThisWorkbook',
        type: VBAModuleType.Document,
        code: `' Code could not be fully extracted for ThisWorkbook`,
        extractionSuccess: false
      });
      
      // Check for worksheet modules
      if (workbook.SheetNames && workbook.SheetNames.length > 0) {
        for (const sheetName of workbook.SheetNames) {
          modules.push({
            name: sheetName,
            type: VBAModuleType.Document,
            code: `' Code could not be fully extracted for worksheet: ${sheetName}`,
            extractionSuccess: false
          });
        }
      }
      
      logger(`Added ${modules.length} potential modules based on workbook structure`, 'info');
    }
    
    // Sort modules by type and name
    modules.sort((a, b) => {
      if (a.type !== b.type) {
        return a.type - b.type;
      }
      return a.name.localeCompare(b.name);
    });
    
    logger(`Alternative extraction found ${modules.length} modules`, modules.length > 0 ? 'success' : 'warning');
    return modules;
  } catch (error) {
    logger(`Error in alternative VBA extraction: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return [];
  }
} 