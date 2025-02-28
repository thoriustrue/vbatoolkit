import { VBAModule, VBAModuleType } from './types';
import { LoggerCallback } from '../../types';
import * as XLSX from 'xlsx';

/**
 * Extracts VBA modules from a workbook
 * @param workbook The XLSX workbook
 * @param logger Callback function for logging messages
 * @returns Array of VBA modules
 */
export function extractVBAModulesFromWorkbook(
  workbook: XLSX.WorkBook,
  logger: LoggerCallback
): VBAModule[] {
  try {
    logger('Attempting to extract VBA modules from workbook...', 'info');
    
    // Check if workbook has VBA code
    if (!workbook.Workbook || !workbook.Workbook.VBAProject) {
      logger('No VBA project found in workbook', 'warning');
      return [];
    }
    
    const modules: VBAModule[] = [];
    const vbaProject = workbook.Workbook.VBAProject;
    
    // Access the VBA project structure
    logger('Accessing VBA project structure...', 'info');
    
    // Extract modules from the VBA project
    if (vbaProject.modules && Array.isArray(vbaProject.modules)) {
      logger(`Found ${vbaProject.modules.length} modules in VBA project`, 'info');
      
      for (const module of vbaProject.modules) {
        if (!module || !module.name) continue;
        
        const name = module.name;
        let type = VBAModuleType.Unknown;
        let code = module.code || '';
        
        // Determine module type based on name and content
        if (name.toLowerCase() === 'thisdocument' || name.toLowerCase() === 'thisworkbook') {
          type = VBAModuleType.Document;
        } else if (name.toLowerCase().startsWith('sheet') || name.toLowerCase().startsWith('worksheet')) {
          type = VBAModuleType.Document;
        } else if (name.toLowerCase().includes('class') || code.toLowerCase().includes('attribute vb_creatable')) {
          type = VBAModuleType.Class;
        } else if (name.toLowerCase().includes('form') || code.toLowerCase().includes('begin vb.form')) {
          type = VBAModuleType.Form;
        } else {
          type = VBAModuleType.Standard;
        }
        
        modules.push({
          name,
          type,
          code: code || `' Code could not be fully extracted for module: ${name}`,
          extractionSuccess: !!code
        });
        
        logger(`Extracted module: ${name} (${VBAModuleType[type]})`, 'info');
      }
    } else {
      logger('No modules found in VBA project structure', 'warning');
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
    
    logger(`Successfully extracted ${modules.length} VBA modules`, 'success');
    return modules;
  } catch (error) {
    logger(`Error extracting VBA modules: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return [];
  }
}

/**
 * Extracts VBA modules from binary data
 * @param data The binary data of the VBA project
 * @param logger Callback function for logging messages
 * @returns Array of VBA modules
 */
export function extractVBAModulesFromBinary(
  data: Uint8Array,
  logger: LoggerCallback
): VBAModule[] {
  try {
    logger('Attempting to extract VBA modules from binary data...', 'info');
    
    const modules: VBAModule[] = [];
    const textDecoder = new TextDecoder('utf-8');
    
    // Convert binary data to string for analysis
    const content = textDecoder.decode(data);
    
    // Find module names and types in the binary content
    // Module names are often stored with a prefix like "MODULE=" or "CLASS="
    const moduleMatches = content.match(/MODULE=([^\r\n]+)/g) || [];
    const classMatches = content.match(/CLASS=([^\r\n]+)/g) || [];
    const documentMatches = content.match(/DOCUMENT=([^\r\n]+)/g) || [];
    
    // Process module matches
    for (const match of moduleMatches) {
      const name = match.substring(7).trim();
      if (!name) continue;
      
      modules.push({
        name,
        type: VBAModuleType.Standard,
        code: `' Code could not be fully extracted for module: ${name}`,
        extractionSuccess: false
      });
      
      logger(`Found standard module: ${name}`, 'info');
    }
    
    // Process class matches
    for (const match of classMatches) {
      const name = match.substring(6).trim();
      if (!name) continue;
      
      modules.push({
        name,
        type: VBAModuleType.Class,
        code: `' Code could not be fully extracted for class module: ${name}`,
        extractionSuccess: false
      });
      
      logger(`Found class module: ${name}`, 'info');
    }
    
    // Process document matches
    for (const match of documentMatches) {
      const name = match.substring(9).trim();
      if (!name) continue;
      
      modules.push({
        name,
        type: VBAModuleType.Document,
        code: `' Code could not be fully extracted for document module: ${name}`,
        extractionSuccess: false
      });
      
      logger(`Found document module: ${name}`, 'info');
    }
    
    // Look for form modules
    const formMatches = content.match(/Begin VB\.Form ([^\r\n]+)/g) || [];
    for (const match of formMatches) {
      const name = match.substring(14).trim();
      if (!name) continue;
      
      modules.push({
        name,
        type: VBAModuleType.Form,
        code: `' Code could not be fully extracted for form: ${name}`,
        extractionSuccess: false
      });
      
      logger(`Found form module: ${name}`, 'info');
    }
    
    // If no modules found, try alternative approach
    if (modules.length === 0) {
      logger('No modules found with standard patterns, trying alternative approach...', 'info');
      
      // Look for attribute VB_Name which often indicates module names
      const nameMatches = content.match(/Attribute VB_Name = "([^"]+)"/g) || [];
      for (const match of nameMatches) {
        const name = match.match(/"([^"]+)"/)?.[1] || '';
        if (!name) continue;
        
        // Try to determine module type
        let type = VBAModuleType.Unknown;
        
        if (name.toLowerCase() === 'thisdocument' || name.toLowerCase() === 'thisworkbook') {
          type = VBAModuleType.Document;
        } else if (name.toLowerCase().startsWith('sheet') || name.toLowerCase().startsWith('worksheet')) {
          type = VBAModuleType.Document;
        } else if (content.includes(`Attribute VB_Name = "${name}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False`)) {
          type = VBAModuleType.Class;
        } else if (content.includes(`Attribute VB_Name = "${name}"\r\nBegin VB.Form`)) {
          type = VBAModuleType.Form;
        } else {
          type = VBAModuleType.Standard;
        }
        
        modules.push({
          name,
          type,
          code: `' Code could not be fully extracted for module: ${name}`,
          extractionSuccess: false
        });
        
        logger(`Found module using attributes: ${name} (${VBAModuleType[type]})`, 'info');
      }
    }
    
    // Sort modules by type and name
    modules.sort((a, b) => {
      if (a.type !== b.type) {
        return a.type - b.type;
      }
      return a.name.localeCompare(b.name);
    });
    
    logger(`Extracted ${modules.length} module names from binary data`, modules.length > 0 ? 'success' : 'warning');
    return modules;
  } catch (error) {
    logger(`Error extracting VBA modules from binary: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return [];
  }
}

/**
 * Attempts to extract code from VBA modules
 * @param data The binary data of the VBA project
 * @param modules The array of VBA modules to populate with code
 * @param logger Callback function for logging messages
 * @returns Updated array of VBA modules with code
 */
export function extractCodeFromModules(
  data: Uint8Array,
  modules: VBAModule[],
  logger: LoggerCallback
): VBAModule[] {
  try {
    logger('Attempting to extract code from VBA modules...', 'info');
    
    const textDecoder = new TextDecoder('utf-8');
    const content = textDecoder.decode(data);
    
    let extractionCount = 0;
    
    // For each module, try to find its code
    for (let i = 0; i < modules.length; i++) {
      const module = modules[i];
      
      // Look for code between module markers
      const moduleStart = content.indexOf(`Attribute VB_Name = "${module.name}"`);
      if (moduleStart >= 0) {
        // Find the next module start or end of file
        let nextModuleStart = content.length;
        for (const otherModule of modules) {
          if (otherModule.name === module.name) continue;
          
          const otherStart = content.indexOf(`Attribute VB_Name = "${otherModule.name}"`);
          if (otherStart > moduleStart && otherStart < nextModuleStart) {
            nextModuleStart = otherStart;
          }
        }
        
        // Extract the code between this module and the next
        let moduleCode = content.substring(moduleStart, nextModuleStart).trim();
        
        // Clean up the code
        moduleCode = cleanModuleCode(moduleCode);
        
        if (moduleCode) {
          modules[i].code = moduleCode;
          modules[i].extractionSuccess = true;
          extractionCount++;
          logger(`Successfully extracted code for module: ${module.name}`, 'info');
        }
      }
    }
    
    logger(`Successfully extracted code for ${extractionCount} out of ${modules.length} modules`, 
      extractionCount > 0 ? 'success' : 'warning');
    
    return modules;
  } catch (error) {
    logger(`Error extracting code from modules: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return modules;
  }
}

/**
 * Cleans up extracted module code
 * @param code The raw module code
 * @returns Cleaned module code
 */
function cleanModuleCode(code: string): string {
  // Remove binary artifacts
  code = code.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  
  // Extract only the VBA code part
  const codeStart = code.indexOf('Attribute VB_Name =');
  if (codeStart >= 0) {
    // Find the end of attributes section
    let codeBody = code.substring(codeStart);
    
    // Look for the first actual code line after attributes
    const attributeEnd = codeBody.search(/(?:^|\r\n)(?!Attribute VB_)/m);
    if (attributeEnd > 0) {
      codeBody = codeBody.substring(attributeEnd).trim();
      return codeBody;
    }
  }
  
  return code;
} 