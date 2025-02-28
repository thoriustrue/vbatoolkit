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
        
        // Check for required namespaces
        const workbook = doc.getElementsByTagName('workbook')[0];
        if (workbook) {
          // Check for required namespaces
          const requiredNamespaces = [
            { prefix: 'xmlns', uri: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' },
            { prefix: 'xmlns:r', uri: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }
          ];
          
          for (const ns of requiredNamespaces) {
            if (!workbook.hasAttribute(ns.prefix) || workbook.getAttribute(ns.prefix) !== ns.uri) {
              workbook.setAttribute(ns.prefix, ns.uri);
              modified = true;
              logger(`Added missing namespace ${ns.prefix}="${ns.uri}" to workbook.xml`, 'info');
            }
          }
          
          // Ensure workbookPr exists
          let workbookPr = doc.getElementsByTagName('workbookPr')[0];
          if (!workbookPr) {
            workbookPr = doc.createElement('workbookPr');
            workbookPr.setAttribute('codeName', 'ThisWorkbook');
            
            // Insert after fileVersion if it exists, otherwise as first child
            const fileVersion = doc.getElementsByTagName('fileVersion')[0];
            if (fileVersion && fileVersion.parentNode) {
              fileVersion.parentNode.insertBefore(workbookPr, fileVersion.nextSibling);
            } else {
              workbook.insertBefore(workbookPr, workbook.firstChild);
            }
            
            modified = true;
            logger('Added missing workbookPr element', 'info');
          } else {
            // Ensure codeName attribute exists
            if (!workbookPr.hasAttribute('codeName')) {
              workbookPr.setAttribute('codeName', 'ThisWorkbook');
              modified = true;
              logger('Added missing codeName attribute to workbookPr', 'info');
            }
          }
        }
        
        // Check for duplicate sheet IDs
        const sheets = doc.getElementsByTagName('sheet');
        const usedIds = new Set();
        const usedNames = new Set();
        
        for (let i = 0; i < sheets.length; i++) {
          let sheetId = sheets[i].getAttribute('sheetId');
          let name = sheets[i].getAttribute('name');
          let rId = sheets[i].getAttribute('r:id');
          
          // Fix missing r:id attribute
          if (!rId) {
            sheets[i].setAttribute('r:id', `rId${i + 1}`);
            modified = true;
            logger(`Added missing r:id attribute to sheet ${name}`, 'info');
          }
          
          if (usedIds.has(sheetId)) {
            // Duplicate ID found, generate a new one
            const newId = (i + 1).toString();
            sheets[i].setAttribute('sheetId', newId);
            modified = true;
            logger(`Fixed duplicate sheetId for ${name}`, 'info');
          }
          
          if (usedNames.has(name)) {
            // Duplicate name found, generate a new one
            const newName = `Sheet${i + 1}`;
            sheets[i].setAttribute('name', newName);
            modified = true;
            logger(`Fixed duplicate sheet name: ${name} -> ${newName}`, 'info');
          }
          
          usedIds.add(sheets[i].getAttribute('sheetId'));
          usedNames.add(sheets[i].getAttribute('name'));
        }
        
        // Ensure sheets element exists and has at least one sheet
        const sheetsElement = doc.getElementsByTagName('sheets')[0];
        if (!sheetsElement || sheetsElement.childNodes.length === 0) {
          // If no sheets element, create one with a default sheet
          if (!sheetsElement) {
            const newSheetsElement = doc.createElement('sheets');
            const newSheet = doc.createElement('sheet');
            newSheet.setAttribute('name', 'Sheet1');
            newSheet.setAttribute('sheetId', '1');
            newSheet.setAttribute('r:id', 'rId1');
            newSheetsElement.appendChild(newSheet);
            
            // Insert after workbookPr
            const workbookPr = doc.getElementsByTagName('workbookPr')[0];
            if (workbookPr && workbookPr.parentNode) {
              workbookPr.parentNode.insertBefore(newSheetsElement, workbookPr.nextSibling);
            } else {
              workbook.appendChild(newSheetsElement);
            }
            
            modified = true;
            logger('Added missing sheets element with default sheet', 'info');
          } 
          // If sheets element exists but has no children, add a default sheet
          else if (sheetsElement.childNodes.length === 0) {
            const newSheet = doc.createElement('sheet');
            newSheet.setAttribute('name', 'Sheet1');
            newSheet.setAttribute('sheetId', '1');
            newSheet.setAttribute('r:id', 'rId1');
            sheetsElement.appendChild(newSheet);
            
            modified = true;
            logger('Added default sheet to empty sheets element', 'info');
          }
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
              const parentNode = cells[i].parentNode;
              if (parentNode) {
                parentNode.removeChild(cells[i]);
                modified = true;
              }
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
    
    // 6. Preserve all original files and directories
    // This ensures we don't lose any components during processing
    const allFiles = Object.keys(zip.files);
    
    // Check for missing critical directories
    const criticalDirs = [
      'xl/',
      'xl/worksheets/',
      'xl/_rels/',
      '_rels/'
    ];
    
    for (const dir of criticalDirs) {
      if (!zip.files[dir]) {
        // Create the directory if it's missing
        zip.folder(dir);
        logger(`Created missing directory: ${dir}`, 'info');
      }
    }
    
    // 7. Fix content types if needed
    const contentTypesFile = zip.file('[Content_Types].xml');
    if (contentTypesFile) {
      let contentTypes = await contentTypesFile.async('string');
      let modified = false;
      
      // Check for required content types
      const requiredTypes = [
        { part: 'workbook.xml', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
        { part: 'vbaProject.bin', type: 'application/vnd.ms-office.vbaProject' },
        { extension: 'rels', type: 'application/vnd.openxmlformats-package.relationships+xml' },
        { extension: 'xml', type: 'application/xml' }
      ];
      
      for (const req of requiredTypes) {
        if (req.part && !contentTypes.includes(`PartName="/${req.part.startsWith('xl/') ? '' : 'xl/'}${req.part}"`)) {
          contentTypes = contentTypes.replace(
            /<Types[^>]*>/,
            `$&\n  <Override PartName="/${req.part.startsWith('xl/') ? '' : 'xl/'}${req.part}" ContentType="${req.type}"/>`
          );
          modified = true;
          logger(`Added missing content type for ${req.part}`, 'info');
        } else if (req.extension && !contentTypes.includes(`Extension="${req.extension}"`)) {
          contentTypes = contentTypes.replace(
            /<Types[^>]*>/,
            `$&\n  <Default Extension="${req.extension}" ContentType="${req.type}"/>`
          );
          modified = true;
          logger(`Added missing content type for extension .${req.extension}`, 'info');
        }
      }
      
      // Add worksheet content types if missing
      for (const file of allFiles) {
        if (file.startsWith('xl/worksheets/sheet') && file.endsWith('.xml')) {
          const sheetPart = file.substring(3); // Remove 'xl/' prefix
          if (!contentTypes.includes(`PartName="/${sheetPart}"`)) {
            contentTypes = contentTypes.replace(
              /<Types[^>]*>/,
              `$&\n  <Override PartName="/${sheetPart}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
            );
            modified = true;
            logger(`Added missing content type for ${file}`, 'info');
          }
        }
      }
      
      if (modified) {
        zip.file('[Content_Types].xml', contentTypes);
        logger('Updated content types file with missing entries', 'success');
      }
    }
    
    // 8. Fix workbook relationships if needed
    const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (workbookRelsFile) {
      let workbookRels = await workbookRelsFile.async('string');
      let modified = false;
      
      // Check for VBA project relationship
      if (!workbookRels.includes('vnd.ms-office.vbaProject') && zip.file('xl/vbaProject.bin')) {
        // Generate a unique rId
        let maxRid = 0;
        const ridMatches = workbookRels.match(/Id="rId(\d+)"/g);
        if (ridMatches) {
          for (const match of ridMatches) {
            const ridNum = parseInt(match.replace(/Id="rId(\d+)"/, '$1'));
            if (ridNum > maxRid) {
              maxRid = ridNum;
            }
          }
        }
        
        const newRid = `rId${maxRid + 1}`;
        workbookRels = workbookRels.replace(
          /<Relationships[^>]*>/,
          `$&\n  <Relationship Id="${newRid}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`
        );
        modified = true;
        logger('Added missing VBA project relationship', 'info');
      }
      
      // Check for worksheet relationships
      for (const file of allFiles) {
        if (file.startsWith('xl/worksheets/sheet') && file.endsWith('.xml')) {
          const sheetName = file.replace('xl/worksheets/', '');
          if (!workbookRels.includes(`Target="worksheets/${sheetName}"`)) {
            // Generate a unique rId
            let maxRid = 0;
            const ridMatches = workbookRels.match(/Id="rId(\d+)"/g);
            if (ridMatches) {
              for (const match of ridMatches) {
                const ridNum = parseInt(match.replace(/Id="rId(\d+)"/, '$1'));
                if (ridNum > maxRid) {
                  maxRid = ridNum;
                }
              }
            }
            
            const newRid = `rId${maxRid + 1}`;
            workbookRels = workbookRels.replace(
              /<Relationships[^>]*>/,
              `$&\n  <Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetName}"/>`
            );
            modified = true;
            logger(`Added missing relationship for ${sheetName}`, 'info');
          }
        }
      }
      
      if (modified) {
        zip.file('xl/_rels/workbook.xml.rels', workbookRels);
        logger('Updated workbook relationships file with missing entries', 'success');
      }
    }
    
    // 9. Remove any potentially corrupted files
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
    logger(`Error during file integrity fix: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return zip;
  }
} 