import JSZip from 'jszip';
import { LoggerCallback } from './types';

/**
 * Removes all sheet-level protections from Excel files
 */
export async function removeSheetProtections(
  zip: JSZip,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    logger('Removing sheet-level protections...', 'info');
    let protectionsRemoved = false;
    
    // Get all worksheet files
    const worksheetFiles = Object.keys(zip.files).filter(
      path => path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')
    );
    
    if (worksheetFiles.length === 0) {
      logger('No worksheet files found', 'warning');
      return false;
    }
    
    logger(`Found ${worksheetFiles.length} worksheets to process`, 'info');
    
    for (const worksheetPath of worksheetFiles) {
      const worksheet = zip.file(worksheetPath);
      if (worksheet) {
        let content = await worksheet.async('string');
        const originalContent = content;
        
        // 1. Remove sheetProtection elements
        if (content.includes('<sheetProtection')) {
          content = content.replace(/<sheetProtection[^>]*\/>/g, '');
          content = content.replace(/<sheetProtection[^>]*>.*?<\/sheetProtection>/gs, '');
          logger(`Removed sheetProtection from ${worksheetPath}`, 'info');
        }
        
        // 2. Remove protectedRanges elements
        if (content.includes('<protectedRanges>')) {
          content = content.replace(/<protectedRanges>.*?<\/protectedRanges>/gs, '');
          logger(`Removed protectedRanges from ${worksheetPath}`, 'info');
        }
        
        // 3. Remove legacy protection attributes
        if (content.includes('protected=')) {
          content = content.replace(/protected="1"/g, 'protected="0"');
          logger(`Removed legacy protection attributes from ${worksheetPath}`, 'info');
        }
        
        // 4. Check if content was modified
        if (content !== originalContent) {
          zip.file(worksheetPath, content);
          protectionsRemoved = true;
        }
      }
    }
    
    // Also check for workbook-level sheet protection
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      let content = await workbookFile.async('string');
      const originalContent = content;
      
      // Remove sheet protection references in workbook
      if (content.includes('<sheetProtection')) {
        content = content.replace(/<sheetProtection[^>]*\/>/g, '');
        zip.file('xl/workbook.xml', content);
        protectionsRemoved = true;
        logger('Removed sheet protection references from workbook.xml', 'info');
      }
    }
    
    if (protectionsRemoved) {
      logger('Successfully removed all sheet protections', 'success');
    } else {
      logger('No sheet protections found to remove', 'info');
    }
    
    return protectionsRemoved;
  } catch (error) {
    logger(`Error removing sheet protections: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
} 