import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';
import * as XLSX from 'xlsx';

/**
 * Extracts VBA modules from a workbook using SheetJS
 * @param workbook The workbook to extract VBA modules from
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
export function extractVBAModulesFromWorkbook(workbook: XLSX.WorkBook, logger: LoggerCallback): VBAModule[] {
  const modules: VBAModule[] = [];
  
  try {
    // Access the VBA project
    if (!workbook.Workbook || !(workbook.Workbook as any).VBAProject) {
      return modules;
    }
    
    // Use type assertion to access VBAProject
    const vbaProject = (workbook.Workbook as any).VBAProject;
    
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
        const moduleType = determineModuleType(moduleCode);
        
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
 * Determines the type of a VBA module based on its code content
 * @param moduleCode The VBA module code
 * @returns The determined module type
 */
function determineModuleType(moduleCode: string): VBAModuleType {
  if (!moduleCode.includes('Attribute VB_Name = "')) {
    return VBAModuleType.Unknown;
  }
  
  if (moduleCode.includes('Attribute VB_Base = "0{')) {
    return VBAModuleType.Form;
  } 
  
  if (moduleCode.includes('Attribute VB_PredeclaredId = True')) {
    if (moduleCode.includes('Attribute VB_Exposed = True')) {
      return VBAModuleType.Document;
    } else {
      return VBAModuleType.Standard;
    }
  } 
  
  if (moduleCode.includes('Attribute VB_GlobalNameSpace = False') && 
      moduleCode.includes('Attribute VB_Creatable = False')) {
    return VBAModuleType.Class;
  } 
  
  return VBAModuleType.Standard;
} 