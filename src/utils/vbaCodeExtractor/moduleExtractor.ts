import { WorkBook } from 'xlsx';
import { LoggerCallback } from '../../types';
import { VBAModule, VBAModuleType } from './types';

/**
 * Extracts VBA modules from a workbook
 * @param workbook The workbook containing VBA
 * @param logger Callback function for logging messages
 * @returns An array of VBA modules
 */
export function extractVBAModulesFromWorkbook(
  workbook: WorkBook,
  logger: LoggerCallback
): VBAModule[] {
  const modules: VBAModule[] = [];
  
  try {
    // Check if the workbook has VBA
    if (!workbook.vbaraw) {
      logger('No VBA code found in this workbook', 'warning');
      return modules;
    }
    
    // Try to access the VBA project
    const vbaProject = workbook.Workbook?.VBAProject;
    if (!vbaProject) {
      logger('VBA project structure not accessible', 'warning');
      return modules;
    }
    
    // Extract modules from the VBA project
    if (vbaProject.modules) {
      for (const [name, content] of Object.entries(vbaProject.modules)) {
        // Determine module type based on name
        let moduleType = VBAModuleType.Standard;
        
        if (name === 'ThisWorkbook') {
          moduleType = VBAModuleType.Document;
        } else if (name.startsWith('Sheet')) {
          moduleType = VBAModuleType.Document;
        } else if (name.includes('UserForm')) {
          moduleType = VBAModuleType.Form;
        } else if (name.includes('Class')) {
          moduleType = VBAModuleType.Class;
        }
        
        modules.push({
          name,
          type: moduleType,
          code: typeof content === 'string' ? content : 'Code could not be extracted'
        });
      }
      
      logger(`Extracted ${modules.length} VBA modules from the workbook`, 'success');
    } else {
      logger('No modules found in the VBA project', 'warning');
    }
    
    return modules;
  } catch (error) {
    logger(`Error extracting VBA modules: ${error instanceof Error ? error.message : String(error)}`, 'error');
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