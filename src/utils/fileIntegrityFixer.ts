import JSZip from 'jszip';
import { LoggerCallback } from './types';

/**
 * Fixes common integrity issues in Excel files after modification
 */
export async function fixFileIntegrity(
  zip: JSZip, 
  logger: LoggerCallback
): Promise<JSZip> {
  try {
    logger('Applying file integrity fixes...', 'info');
    
    // 1. Fix Content_Types.xml
    const contentTypes = zip.file('[Content_Types].xml');
    if (contentTypes) {
      let content = await contentTypes.async('string');
      
      // Ensure all required content types are present
      const requiredTypes = [
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>'
      ];
      
      let modified = false;
      for (const type of requiredTypes) {
        if (!content.includes(type.split('"')[1])) {
          content = content.replace('</Types>', `${type}\n</Types>`);
          modified = true;
        }
      }
      
      if (modified) {
        zip.file('[Content_Types].xml', content);
        logger('Fixed content types', 'info');
      }
    }
    
    // 2. Fix relationship files
    const relFiles = Object.keys(zip.files).filter(path => path.endsWith('.rels'));
    for (const relFile of relFiles) {
      const file = zip.file(relFile);
      if (file) {
        let content = await file.async('string');
        
        // Fix empty Relationships elements
        if (content.includes('<Relationships') && 
            !content.includes('<Relationship') && 
            content.includes('</Relationships>')) {
          // Remove empty Relationships element
          zip.remove(relFile);
          logger(`Removed empty relationship file: ${relFile}`, 'info');
        }
      }
    }
    
    // 3. Fix workbook.xml
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      let content = await workbookFile.async('string');
      
      // Ensure workbook has proper XML structure
      if (!content.includes('<?xml')) {
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + content;
        zip.file('xl/workbook.xml', content);
        logger('Fixed workbook XML declaration', 'info');
      }
    }
    
    logger('File integrity fixes applied', 'success');
    return zip;
  } catch (error) {
    logger(`Error fixing file integrity: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return zip; // Return original zip even if fixes fail
  }
} 