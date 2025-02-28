import JSZip from 'jszip';
import { LoggerCallback } from '../types';
import { DOMParser, XMLSerializer } from 'xmldom';

/**
 * Fixes common integrity issues in Excel files after modification
 */
export async function fixFileIntegrity(
  zip: JSZip, 
  logger: LoggerCallback
): Promise<JSZip> {
  try {
    logger('Applying advanced file integrity fixes...', 'info');
    
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
      
      // Ensure all required part types are present
      const requiredParts = [
        { part: '/xl/workbook.xml', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
        { part: '/xl/vbaProject.bin', type: 'application/vnd.ms-office.vbaProject' }
      ];
      
      for (const req of requiredParts) {
        if (zip.file(req.part.substring(1))) {
          let found = false;
          const overrides = doc.getElementsByTagName('Override');
          for (let i = 0; i < overrides.length; i++) {
            const part = overrides[i].getAttribute('PartName');
            if (part === req.part) {
              found = true;
              break;
            }
          }
          
          if (!found) {
            const newOverride = doc.createElement('Override');
            newOverride.setAttribute('PartName', req.part);
            newOverride.setAttribute('ContentType', req.type);
            types.appendChild(newOverride);
            modified = true;
          }
        }
      }
      
      if (modified) {
        const serializer = new XMLSerializer();
        const newContent = serializer.serializeToString(doc);
        zip.file('[Content_Types].xml', newContent);
        logger('Fixed content types with proper XML parsing', 'info');
      }
    }
    
    // 2. Fix relationship files
    const relFiles = Object.keys(zip.files).filter(path => path.endsWith('.rels'));
    for (const relFile of relFiles) {
      const file = zip.file(relFile);
      if (file) {
        let content = await file.async('string');
        
        try {
          // Parse and fix XML structure
          const parser = new DOMParser();
          const doc = parser.parseFromString(content, 'text/xml');
          
          // Check if Relationships element exists and has children
          const relationships = doc.getElementsByTagName('Relationships')[0];
          if (relationships && relationships.childNodes.length === 0) {
            // Empty relationships file - remove it
            zip.remove(relFile);
            logger(`Removed empty relationship file: ${relFile}`, 'info');
          } else {
            // Fix relationship IDs to ensure uniqueness
            const rels = doc.getElementsByTagName('Relationship');
            const usedIds = new Set();
            let modified = false;
            
            for (let i = 0; i < rels.length; i++) {
              let id = rels[i].getAttribute('Id');
              if (usedIds.has(id)) {
                // Duplicate ID found, generate a new one
                const newId = `rId${i + 100}`; // Use a high number to avoid conflicts
                rels[i].setAttribute('Id', newId);
                modified = true;
              }
              usedIds.add(rels[i].getAttribute('Id'));
            }
            
            if (modified) {
              const serializer = new XMLSerializer();
              const newContent = serializer.serializeToString(doc);
              zip.file(relFile, newContent);
              logger(`Fixed relationship IDs in ${relFile}`, 'info');
            }
          }
        } catch (xmlError) {
          logger(`Error parsing XML in ${relFile}: ${xmlError}`, 'warning');
        }
      }
    }
    
    // 3. Fix workbook.xml
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      let content = await workbookFile.async('string');
      
      try {
        // Parse and fix XML structure
        const parser = new DOMParser();
        const doc = parser.parseFromString(content, 'text/xml');
        
        // Ensure workbook has proper XML structure
        let modified = false;
        
        // Check for duplicate sheet IDs
        const sheets = doc.getElementsByTagName('sheet');
        const usedIds = new Set();
        const usedNames = new Set();
        
        for (let i = 0; i < sheets.length; i++) {
          let sheetId = sheets[i].getAttribute('sheetId');
          let name = sheets[i].getAttribute('name');
          
          if (usedIds.has(sheetId)) {
            // Duplicate ID found, generate a new one
            const newId = (i + 1).toString();
            sheets[i].setAttribute('sheetId', newId);
            modified = true;
          }
          
          if (usedNames.has(name)) {
            // Duplicate name found, generate a new one
            const newName = `Sheet${i + 1}`;
            sheets[i].setAttribute('name', newName);
            modified = true;
          }
          
          usedIds.add(sheets[i].getAttribute('sheetId'));
          usedNames.add(sheets[i].getAttribute('name'));
        }
        
        if (modified) {
          const serializer = new XMLSerializer();
          const newContent = serializer.serializeToString(doc);
          zip.file('xl/workbook.xml', newContent);
          logger('Fixed workbook XML structure', 'info');
        }
      } catch (xmlError) {
        logger(`Error parsing workbook XML: ${xmlError}`, 'warning');
      }
    }
    
    // 4. Fix worksheet files
    const worksheetFiles = Object.keys(zip.files).filter(
      path => path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')
    );
    
    for (const worksheetPath of worksheetFiles) {
      const worksheet = zip.file(worksheetPath);
      if (worksheet) {
        let content = await worksheet.async('string');
        
        try {
          // Parse and fix XML structure
          const parser = new DOMParser();
          const doc = parser.parseFromString(content, 'text/xml');
          
          // Check for and fix common worksheet issues
          let modified = false;
          
          // Fix missing dimension element
          const dimensions = doc.getElementsByTagName('dimension');
          if (dimensions.length === 0) {
            const worksheet = doc.getElementsByTagName('worksheet')[0];
            const sheetData = doc.getElementsByTagName('sheetData')[0];
            
            if (worksheet && sheetData) {
              const dimension = doc.createElement('dimension');
              dimension.setAttribute('ref', 'A1');
              worksheet.insertBefore(dimension, sheetData);
              modified = true;
            }
          }
          
          // Remove invalid cell references
          const cells = doc.getElementsByTagName('c');
          for (let i = cells.length - 1; i >= 0; i--) {
            const r = cells[i].getAttribute('r');
            if (!r || !/^[A-Z]+[0-9]+$/.test(r)) {
              cells[i].parentNode.removeChild(cells[i]);
              modified = true;
            }
          }
          
          if (modified) {
            const serializer = new XMLSerializer();
            const newContent = serializer.serializeToString(doc);
            zip.file(worksheetPath, newContent);
            logger(`Fixed worksheet XML in ${worksheetPath}`, 'info');
          }
        } catch (xmlError) {
          logger(`Error parsing worksheet XML in ${worksheetPath}: ${xmlError}`, 'warning');
        }
      }
    }
    
    // 5. Fix VBA project binary structure
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
    
    // 6. Remove any potentially corrupted files
    const knownCorruptPatterns = [
      'xl/ctrlProps/',
      'xl/activeX/'
      // Removed 'xl/drawings/vmlDrawing' as these files are often needed for shapes and buttons
    ];
    
    // Only remove files that are actually corrupted
    for (const pattern of knownCorruptPatterns) {
      const matchingFiles = Object.keys(zip.files).filter(path => path.includes(pattern));
      for (const file of matchingFiles) {
        try {
          // Try to read the file first to check if it's actually corrupted
          const fileObj = zip.file(file);
          if (!fileObj) {
            logger(`File not found: ${file}`, 'info');
            continue;
          }
          
          const content = await fileObj.async('uint8array');
          if (content.length === 0) {
            // Empty file, safe to remove
            zip.remove(file);
            logger(`Removed empty file: ${file}`, 'info');
          } else if (pattern === 'xl/ctrlProps/' && content.length < 10) {
            // Control properties files should have a minimum size
            zip.remove(file);
            logger(`Removed potentially corrupted file: ${file}`, 'info');
          }
        } catch (err) {
          // If we can't read the file, it's likely corrupted
          zip.remove(file);
          logger(`Removed unreadable file: ${file}`, 'info');
        }
      }
    }
    
    logger('Advanced file integrity fixes applied', 'success');
    return zip;
  } catch (error) {
    logger(`Error fixing file integrity: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return zip; // Return original zip even if fixes fail
  }
} 