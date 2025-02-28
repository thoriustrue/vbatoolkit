import JSZip from 'jszip';
import { LoggerCallback } from './types';
import { DOMParser, XMLSerializer } from 'xmldom';

/**
 * Fixes common integrity issues in Excel files after modification
 */
export async function fixFileIntegrity(
  zip: JSZip, 
  logger: LoggerCallback
): Promise<JSZip> {
  try {
    logger('Applying targeted file integrity fixes...', 'info');
    
    // 1. Fix Content_Types.xml
    const contentTypes = zip.file('[Content_Types].xml');
    if (contentTypes) {
      let content = await contentTypes.async('string');
      
      // Ensure all required content types are present
      const requiredTypes = [
        { ext: 'bin', type: 'application/vnd.ms-office.vbaProject' },
        { ext: 'rels', type: 'application/vnd.openxmlformats-package.relationships+xml' },
        { ext: 'xml', type: 'application/xml' },
        { ext: 'vml', type: 'application/vnd.openxmlformats-officedocument.vmlDrawing' }
      ];
      
      let modified = false;
      const parser = new DOMParser();
      const doc = parser.parseFromString(content, 'text/xml');
      const types = doc.getElementsByTagName('Types')[0];
      
      for (const req of requiredTypes) {
        let found = false;
        const defaults = doc.getElementsByTagName('Default');
        for (let i = 0; i < defaults.length; i++) {
          const ext = defaults[i].getAttribute('Extension');
          if (ext === req.ext) {
            found = true;
            break;
          }
        }
        
        if (!found) {
          const newDefault = doc.createElement('Default');
          newDefault.setAttribute('Extension', req.ext);
          newDefault.setAttribute('ContentType', req.type);
          types.appendChild(newDefault);
          modified = true;
        }
      }
      
      if (modified) {
        const serializer = new XMLSerializer();
        content = serializer.serializeToString(doc);
        zip.file('[Content_Types].xml', content);
        logger('Fixed content types', 'info');
      }
    }
    
    // 2. Fix workbook.xml
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
    
    // 3. Fix VBA project binary structure
    const vbaProject = zip.file('xl/vbaProject.bin');
    if (vbaProject) {
      const vbaContent = await vbaProject.async('uint8array');
      
      // Check VBA project signature
      if (vbaContent.length > 8) {
        const signature = vbaContent.slice(0, 2);
        if (signature[0] !== 0xCC || signature[1] !== 0x61) {
          // Invalid signature, try to fix it
          const fixedVba = new Uint8Array(vbaContent.length);
          fixedVba.set(vbaContent);
          fixedVba[0] = 0xCC;
          fixedVba[1] = 0x61;
          
          // Recalculate checksum
          let checksum = 0;
          for (let i = 8; i < fixedVba.length; i++) {
            checksum += fixedVba[i];
            checksum &= 0xFFFFFFFF;
          }
          
          const view = new DataView(fixedVba.buffer);
          view.setUint32(4, checksum, true);
          
          zip.file('xl/vbaProject.bin', fixedVba);
          logger('Fixed VBA project signature and checksum', 'info');
        }
      }
    }
    
    // 4. Only remove specific problematic files that are known to cause issues
    // Instead of removing all ctrlProps and vmlDrawing files, only remove ones with specific issues
    const filesToCheck = Object.keys(zip.files).filter(path => 
      path.includes('xl/ctrlProps/') || 
      path.includes('xl/drawings/vmlDrawing')
    );
    
    let removedCount = 0;
    for (const filePath of filesToCheck) {
      try {
        // Only check XML files
        if (filePath.endsWith('.xml')) {
          const fileContent = await zip.file(filePath)?.async('string');
          if (fileContent) {
            // Only remove if the file has actual corruption markers
            if (
              fileContent.includes('<?mso-application progid="Excel.Sheet"?>') || // Invalid header
              fileContent.includes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?xml') || // Double XML declaration
              fileContent.includes('<corrupt>') || // Explicit corruption marker
              !fileContent.trim().startsWith('<?xml') // Missing XML declaration
            ) {
              zip.remove(filePath);
              logger(`Removed corrupted file: ${filePath}`, 'info');
              removedCount++;
            }
          }
        }
      } catch (error) {
        // If we can't even read the file, it's likely corrupted
        zip.remove(filePath);
        logger(`Removed unreadable file: ${filePath}`, 'info');
        removedCount++;
      }
    }
    
    if (removedCount > 0) {
      logger(`Removed ${removedCount} corrupted files`, 'info');
    } else {
      logger('No corrupted files found', 'info');
    }
    
    logger('Targeted file integrity fixes applied', 'success');
    return zip;
  } catch (error) {
    logger(`Error fixing file integrity: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return zip; // Return original zip even if fixes fail
  }
} 