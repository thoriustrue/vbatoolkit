import JSZip from 'jszip';
import { LoggerCallback } from '../types';

/**
 * Removes sheet protections from all worksheets in the workbook
 * @param zip The JSZip instance containing the Excel file
 * @param logger Callback function for logging messages
 */
export async function removeSheetProtections(
  zip: JSZip,
  logger: LoggerCallback
): Promise<void> {
  try {
    logger('Checking for sheet-level protections...', 'info');
    
    // Find all worksheet files
    const worksheetFiles = Object.keys(zip.files).filter(
      (filename) => filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')
    );
    
    if (worksheetFiles.length === 0) {
      logger('No worksheets found in the workbook', 'warning');
      return;
    }
    
    logger(`Found ${worksheetFiles.length} worksheets to check for protections`, 'info');
    
    let protectionsRemoved = 0;
    
    // Process each worksheet
    for (const worksheetPath of worksheetFiles) {
      const file = zip.file(worksheetPath);
      if (!file) continue;
      
      const content = await file.async('string');
      
      // Check if the worksheet has protection
      if (content.includes('<sheetProtection ')) {
        // Instead of removing the protection tag, comment it out to preserve XML structure
        const modifiedContent = content.replace(
          /(<sheetProtection [^>]*>)/g,
          '<!-- $1 -->'
        );
        
        // Update the worksheet file
        zip.file(worksheetPath, modifiedContent);
        
        const sheetName = worksheetPath.split('/').pop()?.replace('.xml', '') || 'Unknown';
        logger(`Removed protection from worksheet: ${sheetName}`, 'success');
        protectionsRemoved++;
      }
    }
    
    if (protectionsRemoved > 0) {
      logger(`Successfully removed protections from ${protectionsRemoved} worksheets`, 'success');
    } else {
      logger('No sheet protections found in this workbook', 'info');
    }
  } catch (error) {
    logger(`Error removing sheet protections: ${error instanceof Error ? error.message : String(error)}`, 'error');
  }
}

/**
 * Checks if a workbook has any protected sheets
 * @param zip The JSZip instance containing the Excel file
 * @returns Promise resolving to true if protected sheets are found, false otherwise
 */
export async function hasProtectedSheets(zip: JSZip): Promise<boolean> {
  try {
    // Find all worksheet files
    const worksheetFiles = Object.keys(zip.files).filter(
      (filename) => filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')
    );
    
    // Check each worksheet for protection
    for (const worksheetPath of worksheetFiles) {
      const file = zip.file(worksheetPath);
      if (!file) continue;
      
      const content = await file.async('string');
      
      // Check if the worksheet has protection
      if (content.includes('<sheetProtection ')) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.error('Error checking for protected sheets:', error);
    return false;
  }
} 