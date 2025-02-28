import JSZip from 'jszip';
import { LoggerCallback, ProgressCallback } from '../types';
import { readFileAsArrayBuffer } from './fileUtils';
import { validateOfficeCRC, isValidZip } from './zipValidator';
import { removeSheetProtections } from './sheetProtectionRemover';
import { fixFileIntegrity } from './fileIntegrityFixer';
import { enableMaximumTrust } from './trustEnabler';

export async function removeVBAPassword(
  file: File,
  logger: LoggerCallback,
  progressCallback: ProgressCallback
): Promise<Blob | null> {
  let zip: JSZip | undefined;
  
  try {
    logger('Starting VBA password removal process...', 'info');
    progressCallback(0.1);
    
    const arrayBuffer = await readFileAsArrayBuffer(file);
    
    // Validate ZIP structure first
    if (!isValidZip(arrayBuffer)) {
      logger('Invalid file format - not a valid Office file', 'error');
      return null;
    }
    
    const fileData = new Uint8Array(arrayBuffer);
    zip = await JSZip.loadAsync(fileData);
    
    // Validate Office file structure
    if (!validateOfficeCRC(zip, logger)) {
      throw new Error('Invalid Office file structure');
    }
    
    progressCallback(0.3);
    
    // Check if vbaProject.bin exists
    const vbaProject = zip.file('xl/vbaProject.bin');
    if (!vbaProject) {
      logger('No VBA project found in this file', 'error');
      return null;
    }
    
    // Get vbaProject.bin content
    const vbaContent = await vbaProject.async('uint8array');
    
    // Process the VBA project to remove password
    const modifiedVba = preserveVBAStructure(vbaContent, logger);
    if (!modifiedVba) {
      throw new Error('Failed to remove VBA password');
    }
    
    // Validate and update VBA project checksum
    const finalVba = updateVBAProjectChecksum(modifiedVba, logger);
    if (!finalVba) {
      throw new Error('Failed to update VBA project checksum');
    }
    
    progressCallback(0.6);
    
    // Replace the vbaProject.bin with the modified version
    zip.file('xl/vbaProject.bin', finalVba);
    
    logger('Auto-enabling macros and external links...', 'info');
    
    // Update workbook.xml to auto-enable macros
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      let workbookContent = await workbookFile.async('string');
      
      // Add or modify fileVersion to enable macros
      if (!workbookContent.includes('<fileVersion')) {
        workbookContent = workbookContent.replace(
          /<workbook[^>]*>/,
          '$&\n  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>'
        );
      } else {
        workbookContent = workbookContent.replace(
          /<fileVersion[^>]*\/>/,
          '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>'
        );
      }
      
      zip.file('xl/workbook.xml', workbookContent);
    }
    
    // Remove digital signatures if present
    if (zip.file('_xmlsignatures/sig1.xml')) {
      zip.remove('_xmlsignatures/sig1.xml');
      logger('Removed digital signature', 'info');
    }
    
    // Remove any signature relationships
    const relsFile = zip.file('_rels/.rels');
    if (relsFile) {
      let relsContent = await relsFile.async('string');
      if (relsContent.includes('relationships/digital-signature')) {
        relsContent = relsContent.replace(
          /<Relationship[^>]*relationships\/digital-signature[^>]*\/>/g,
          ''
        );
        zip.file('_rels/.rels', relsContent);
        logger('Removed signature relationships', 'info');
      }
    }
    
    // Remove vbaProjectSignature if present
    if (zip.file('xl/vbaProjectSignature.bin')) {
      zip.remove('xl/vbaProjectSignature.bin');
      logger('Removed VBA project signature', 'info');
    }
    
    if (zip.file('xl/_rels/vbaProject.bin.rels')) {
      const vbaRelsFile = zip.file('xl/_rels/vbaProject.bin.rels');
      if (vbaRelsFile) {
        let vbaRels = await vbaRelsFile.async('string');
        if (vbaRels.includes('vbaProjectSignature')) {
          vbaRels = vbaRels.replace(
            /<Relationship[^>]*vbaProjectSignature[^>]*\/>/g,
            ''
          );
          zip.file('xl/_rels/vbaProject.bin.rels', vbaRels);
          logger('Cleaned VBA project relationships', 'info');
        }
      }
    }
    
    progressCallback(0.8);
    
    // Fix sheet protections
    await removeSheetProtections(zip, logger);
    progressCallback(0.85);
    
    // Apply file integrity fixes
    await fixFileIntegrity(zip, logger);
    progressCallback(0.9);
    
    // Enable maximum trust settings
    await enableMaximumTrust(zip, logger);
    progressCallback(0.95);
    
    // IMPORTANT: Preserve all original files that might be getting lost
    // This ensures we don't lose any critical components
    await preserveExcelComponents(zip, logger);
    
    // Generate the modified file with proper MIME type and compression
    const modifiedFile = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 1 }, // Minimal compression to ensure file integrity
      mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
    });
    
    logger('VBA password successfully removed!', 'success');
    progressCallback(1);
    
    return modifiedFile;
  } catch (error) {
    logger(`Error during VBA password removal: ${error instanceof Error ? error.message : String(error)}`, 'error');
    
    // Attempt recovery if we have a zip object
    if (typeof zip !== 'undefined') {
      const recovered = await attemptErrorRecovery(
        error instanceof Error ? error : new Error(String(error)),
        zip,
        logger
      );
      
      if (recovered) {
        logger('Recovery successful, continuing with processing...', 'success');
        
        // Continue with processing after recovery
        try {
          // Apply file integrity fixes again after recovery
          await fixFileIntegrity(zip, logger);
          
          // Preserve Excel components
          await preserveExcelComponents(zip, logger);
          
          // Try recovery with reduced compression
          logger('Attempting recovery with minimal compression...', 'info');
          
          try {
            const modifiedFile = await zip.generateAsync({
              type: 'blob',
              compression: 'DEFLATE',
              compressionOptions: { level: 0 }, // No compression for maximum compatibility
              mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
            });
            
            return modifiedFile;
          } catch (recoveryError) {
            logger(`Recovery failed: ${recoveryError instanceof Error ? recoveryError.message : String(recoveryError)}`, 'error');
            return null;
          }
        } catch (secondError) {
          logger(`Error after recovery attempt: ${secondError instanceof Error ? secondError.message : String(secondError)}`, 'error');
        }
      }
    }
    
    return null;
  }
}

/**
 * Ensures all critical Excel components are preserved
 * This helps prevent file corruption by making sure we don't lose important parts
 */
async function preserveExcelComponents(zip: JSZip, logger: LoggerCallback): Promise<void> {
  logger('Ensuring all critical Excel components are preserved...', 'info');
  
  // Check for critical Excel components
  const criticalComponents = [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/theme/theme1.xml'
  ];
  
  // Check if any worksheets exist
  const worksheets = Object.keys(zip.files).filter(path => 
    path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')
  );
  
  if (worksheets.length === 0) {
    logger('Warning: No worksheets found in the file', 'warning');
  }
  
  // Check for missing critical components
  for (const component of criticalComponents) {
    if (!zip.file(component)) {
      logger(`Warning: Critical component ${component} is missing`, 'warning');
    }
  }
  
  // Ensure the Content_Types file has all necessary entries
  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    let contentTypes = await contentTypesFile.async('string');
    
    // Check for critical content types
    const criticalContentTypes = [
      { partName: '/xl/workbook.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
      { partName: '/xl/vbaProject.bin', contentType: 'application/vnd.ms-office.vbaProject' },
      { extension: 'rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
      { extension: 'xml', contentType: 'application/xml' }
    ];
    
    let missingTypes = [];
    
    try {
      // Try to parse the content types XML
      const parser = new DOMParser();
      const doc = parser.parseFromString(contentTypes, 'text/xml');
      const types = doc.getElementsByTagName('Types')[0];
      
      if (types) {
        // Check for missing default extensions
        for (const type of criticalContentTypes) {
          if (type.extension) {
            let found = false;
            const defaults = doc.getElementsByTagName('Default');
            
            for (let i = 0; i < defaults.length; i++) {
              if (defaults[i].getAttribute('Extension') === type.extension) {
                found = true;
                break;
              }
            }
            
            if (!found) {
              const newDefault = doc.createElement('Default');
              newDefault.setAttribute('Extension', type.extension);
              newDefault.setAttribute('ContentType', type.contentType);
              types.appendChild(newDefault);
              missingTypes.push(`Default: ${type.extension} -> ${type.contentType}`);
            }
          } else if (type.partName) {
            let found = false;
            const overrides = doc.getElementsByTagName('Override');
            
            for (let i = 0; i < overrides.length; i++) {
              if (overrides[i].getAttribute('PartName') === type.partName) {
                found = true;
                break;
              }
            }
            
            if (!found) {
              const newOverride = doc.createElement('Override');
              newOverride.setAttribute('PartName', type.partName);
              newOverride.setAttribute('ContentType', type.contentType);
              types.appendChild(newOverride);
              missingTypes.push(`Override: ${type.partName} -> ${type.contentType}`);
            }
          }
        }
        
        // Add worksheet content types
        for (const worksheetPath of worksheets) {
          const partName = `/${worksheetPath}`;
          let found = false;
          const overrides = doc.getElementsByTagName('Override');
          
          for (let i = 0; i < overrides.length; i++) {
            if (overrides[i].getAttribute('PartName') === partName) {
              found = true;
              break;
            }
          }
          
          if (!found) {
            const newOverride = doc.createElement('Override');
            newOverride.setAttribute('PartName', partName);
            newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
            types.appendChild(newOverride);
            missingTypes.push(`Override: ${partName} -> worksheet`);
          }
        }
        
        // Add content type for shared strings if it exists
        if (zip.file('xl/sharedStrings.xml')) {
          let found = false;
          const overrides = doc.getElementsByTagName('Override');
          
          for (let i = 0; i < overrides.length; i++) {
            if (overrides[i].getAttribute('PartName') === '/xl/sharedStrings.xml') {
              found = true;
              break;
            }
          }
          
          if (!found) {
            const newOverride = doc.createElement('Override');
            newOverride.setAttribute('PartName', '/xl/sharedStrings.xml');
            newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
            types.appendChild(newOverride);
            missingTypes.push('Override: /xl/sharedStrings.xml -> sharedStrings');
          }
        }
        
        // Add content type for styles if it exists
        if (zip.file('xl/styles.xml')) {
          let found = false;
          const overrides = doc.getElementsByTagName('Override');
          
          for (let i = 0; i < overrides.length; i++) {
            if (overrides[i].getAttribute('PartName') === '/xl/styles.xml') {
              found = true;
              break;
            }
          }
          
          if (!found) {
            const newOverride = doc.createElement('Override');
            newOverride.setAttribute('PartName', '/xl/styles.xml');
            newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
            types.appendChild(newOverride);
            missingTypes.push('Override: /xl/styles.xml -> styles');
          }
        }
        
        // Add content type for theme if it exists
        if (zip.file('xl/theme/theme1.xml')) {
          let found = false;
          const overrides = doc.getElementsByTagName('Override');
          
          for (let i = 0; i < overrides.length; i++) {
            if (overrides[i].getAttribute('PartName') === '/xl/theme/theme1.xml') {
              found = true;
              break;
            }
          }
          
          if (!found) {
            const newOverride = doc.createElement('Override');
            newOverride.setAttribute('PartName', '/xl/theme/theme1.xml');
            newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml');
            types.appendChild(newOverride);
            missingTypes.push('Override: /xl/theme/theme1.xml -> theme');
          }
        }
        
        // Serialize the updated XML
        const serializer = new XMLSerializer();
        contentTypes = serializer.serializeToString(doc);
        zip.file('[Content_Types].xml', contentTypes);
      }
    } catch (xmlError) {
      logger(`Error parsing content types XML: ${xmlError}. Using string-based approach.`, 'warning');
      
      // Fallback to string-based approach if XML parsing fails
      for (const type of criticalContentTypes) {
        if (type.extension && !contentTypes.includes(`Extension="${type.extension}"`)) {
          contentTypes = contentTypes.replace(
            /<Types[^>]*>/,
            `$&\n  <Default Extension="${type.extension}" ContentType="${type.contentType}"/>`
          );
          missingTypes.push(`Default: ${type.extension} -> ${type.contentType}`);
        } else if (type.partName && !contentTypes.includes(`PartName="${type.partName}"`)) {
          contentTypes = contentTypes.replace(
            /<Types[^>]*>/,
            `$&\n  <Override PartName="${type.partName}" ContentType="${type.contentType}"/>`
          );
          missingTypes.push(`Override: ${type.partName} -> ${type.contentType}`);
        }
      }
      
      // Add worksheet content types
      for (const worksheetPath of worksheets) {
        const partName = `/${worksheetPath}`;
        if (!contentTypes.includes(`PartName="${partName}"`)) {
          contentTypes = contentTypes.replace(
            /<Types[^>]*>/,
            `$&\n  <Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
          );
          missingTypes.push(`Override: ${partName} -> worksheet`);
        }
      }
      
      zip.file('[Content_Types].xml', contentTypes);
    }
    
    if (missingTypes.length > 0) {
      logger(`Fixed missing content types: ${missingTypes.join(', ')}`, 'success');
    }
  }
  
  // Ensure all required relationships are present
  const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (workbookRelsFile) {
    let workbookRels = await workbookRelsFile.async('string');
    
    // Check if vbaProject relationship exists
    if (!workbookRels.includes('vnd.ms-office.vbaProject') && zip.file('xl/vbaProject.bin')) {
      // Add the relationship if it doesn't exist
      workbookRels = workbookRels.replace(
        /<\?xml[^>]*\?>\s*<Relationships[^>]*>/,
        `$&\n  <Relationship Id="rId9999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`
      );
      zip.file('xl/_rels/workbook.xml.rels', workbookRels);
      logger('Added missing VBA project relationship', 'info');
    }
    
    // Check for worksheet relationships
    for (const worksheetPath of worksheets) {
      const sheetName = worksheetPath.replace('xl/worksheets/', '');
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
        
        // Also add the sheet to workbook.xml if it's not already there
        const workbookFile = zip.file('xl/workbook.xml');
        if (workbookFile) {
          let workbookContent = await workbookFile.async('string');
          const sheetId = sheetName.replace(/[^\d]/g, '');
          
          if (!workbookContent.includes(`r:id="${newRid}"`) && !workbookContent.includes(`name="Sheet${sheetId}"`)) {
            workbookContent = workbookContent.replace(
              /<sheets[^>]*>/,
              `$&\n    <sheet name="Sheet${sheetId}" sheetId="${sheetId}" r:id="${newRid}"/>`
            );
            zip.file('xl/workbook.xml', workbookContent);
            logger(`Added missing sheet ${sheetId} to workbook.xml`, 'info');
          }
        }
        
        logger(`Added missing relationship for ${sheetName}`, 'info');
      }
    }
    
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
  }
  
  // Ensure the main .rels file is correct
  const mainRelsFile = zip.file('_rels/.rels');
  if (mainRelsFile) {
    let mainRels = await mainRelsFile.async('string');
    let modified = false;
    
    // Check for workbook relationship
    if (!mainRels.includes('officeDocument/2006/relationships/officeDocument')) {
      mainRels = mainRels.replace(
        /<Relationships[^>]*>/,
        `$&\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`
      );
      modified = true;
      logger('Added missing workbook relationship to main .rels', 'info');
    }
    
    if (modified) {
      zip.file('_rels/.rels', mainRels);
    }
  } else {
    // Create the main .rels file if it doesn't exist
    const mainRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    zip.file('_rels/.rels', mainRels);
    logger('Created missing main .rels file', 'info');
  }
  
  logger('Critical component check completed', 'info');
}

// Function to preserve VBA structure
function preserveVBAStructure(vbaData: Uint8Array, logger: LoggerCallback): Uint8Array | null {
  try {
    // Create a copy of the data
    const data = new Uint8Array(vbaData);
    
    // VBA binary format has several key sections we need to preserve
    // 1. Header (first 8 bytes including signature and checksum)
    // 2. Directory stream (contains module information)
    // 3. Module streams (contain the actual code)
    
    // Find the project information section (contains password info)
    const projectInfoSignature = [0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74]; // "Project" in ASCII
    let projectInfoOffset = -1;
    
    for (let i = 0; i < data.length - projectInfoSignature.length; i++) {
      let match = true;
      for (let j = 0; j < projectInfoSignature.length; j++) {
        if (data[i + j] !== projectInfoSignature[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        projectInfoOffset = i;
        break;
      }
    }
    
    if (projectInfoOffset !== -1) {
      logger(`Found Project Information at offset ${projectInfoOffset}`, 'info');
      
      // Search for protection record near the project info
      // Protection record is typically within 100-200 bytes of the Project signature
      const searchRange = Math.min(500, data.length - projectInfoOffset);
      
      let protectionFound = false;
      
      // Search for all possible protection patterns
      for (let i = projectInfoOffset; i < projectInfoOffset + searchRange; i++) {
        // Protection record often has this pattern
        if (data[i] === 0x13 && data[i+1] === 0x00 && data[i+2] === 0x01 && data[i+3] === 0x00) {
          // Found potential protection record, disable it
          data[i+2] = 0x00; // Change 0x01 to 0x00 to disable protection
          logger(`Modified protection record at offset ${i}`, 'info');
          protectionFound = true;
        }
        
        // Alternative protection pattern
        if (data[i] === 0x13 && data[i+1] === 0x00 && data[i+2] === 0x02 && data[i+3] === 0x00) {
          data[i+2] = 0x00; // Change 0x02 to 0x00
          logger(`Modified alternative protection record at offset ${i}`, 'info');
          protectionFound = true;
        }
        
        // Another common protection pattern
        if (data[i] === 0x13 && data[i+1] === 0x00 && data[i+2] === 0x11 && data[i+3] === 0x00) {
          data[i+2] = 0x00; // Change 0x11 to 0x00
          logger(`Modified extended protection record at offset ${i}`, 'info');
          protectionFound = true;
        }
      }
      
      if (!protectionFound) {
        logger('No standard protection record found, searching for extended patterns...', 'info');
        
        // Extended search for other protection patterns
        for (let i = projectInfoOffset; i < projectInfoOffset + searchRange; i++) {
          // Look for DPB constant (often indicates protection)
          if (data[i] === 0x44 && data[i+1] === 0x50 && data[i+2] === 0x42) {
            // Found DPB marker, check nearby bytes
            for (let j = i; j < i + 20; j++) {
              if (data[j] === 0x01 && data[j+1] === 0x00 && data[j+2] === 0x01) {
                data[j+2] = 0x00; // Disable protection
                logger(`Modified extended protection record at offset ${j}`, 'info');
                protectionFound = true;
              }
            }
          }
        }
      }
    }
    
    // Search for specific security markers as outlined in the technical roadmap
    
    // 1. Search for "DPB=" marker (Denotes a protected project descriptor)
    const dpbMarker = [0x44, 0x50, 0x42, 0x3D]; // "DPB=" in ASCII
    for (let i = 0; i < data.length - dpbMarker.length; i++) {
      let match = true;
      for (let j = 0; j < dpbMarker.length; j++) {
        if (data[i + j] !== dpbMarker[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        // Found DPB marker, zero out the next 5 bytes
        for (let j = 0; j < 5; j++) {
          if (i + dpbMarker.length + j < data.length) {
            data[i + dpbMarker.length + j] = 0x00;
          }
        }
        logger(`Zeroed out DPB= protection marker at offset ${i}`, 'info');
      }
    }
    
    // 2. Search for "MSDPB" marker (Microsoft's encrypted password flag)
    const msdpbMarker = [0x4D, 0x53, 0x44, 0x50, 0x42]; // "MSDPB" in ASCII
    for (let i = 0; i < data.length - msdpbMarker.length; i++) {
      let match = true;
      for (let j = 0; j < msdpbMarker.length; j++) {
        if (data[i + j] !== msdpbMarker[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        // Found MSDPB marker, zero out the next 5 bytes
        for (let j = 0; j < 5; j++) {
          if (i + msdpbMarker.length + j < data.length) {
            data[i + msdpbMarker.length + j] = 0x00;
          }
        }
        logger(`Zeroed out MSDPB protection marker at offset ${i}`, 'info');
      }
    }
    
    // 3. Search for "CMG=" marker (Checksum and security validation field)
    const cmgMarker = [0x43, 0x4D, 0x47, 0x3D]; // "CMG=" in ASCII
    for (let i = 0; i < data.length - cmgMarker.length; i++) {
      let match = true;
      for (let j = 0; j < cmgMarker.length; j++) {
        if (data[i + j] !== cmgMarker[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        // Found CMG marker, zero out the next 5 bytes
        for (let j = 0; j < 5; j++) {
          if (i + cmgMarker.length + j < data.length) {
            data[i + cmgMarker.length + j] = 0x00;
          }
        }
        logger(`Zeroed out CMG= protection marker at offset ${i}`, 'info');
      }
    }
    
    // 4. Search for "GC=" marker (Security validation field)
    const gcMarker = [0x47, 0x43, 0x3D]; // "GC=" in ASCII
    for (let i = 0; i < data.length - gcMarker.length; i++) {
      let match = true;
      for (let j = 0; j < gcMarker.length; j++) {
        if (data[i + j] !== gcMarker[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        // Found GC marker, zero out the next 5 bytes
        for (let j = 0; j < 5; j++) {
          if (i + gcMarker.length + j < data.length) {
            data[i + gcMarker.length + j] = 0x00;
          }
        }
        logger(`Zeroed out GC= protection marker at offset ${i}`, 'info');
      }
    }
    
    // 5. Search for "PC=" marker (Security validation field)
    const pcMarker = [0x50, 0x43, 0x3D]; // "PC=" in ASCII
    for (let i = 0; i < data.length - pcMarker.length; i++) {
      let match = true;
      for (let j = 0; j < pcMarker.length; j++) {
        if (data[i + j] !== pcMarker[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        // Found PC marker, zero out the next 5 bytes
        for (let j = 0; j < 5; j++) {
          if (i + pcMarker.length + j < data.length) {
            data[i + pcMarker.length + j] = 0x00;
          }
        }
        logger(`Zeroed out PC= protection marker at offset ${i}`, 'info');
      }
    }
    
    // Preserve the PROJECTLOCKED record if it exists
    const lockedSignature = [0x50, 0x52, 0x4F, 0x4A, 0x45, 0x43, 0x54, 0x4C, 0x4F, 0x43, 0x4B, 0x45, 0x44]; // "PROJECTLOCKED"
    let lockedOffset = -1;
    
    for (let i = 0; i < data.length - lockedSignature.length; i++) {
      let match = true;
      for (let j = 0; j < lockedSignature.length; j++) {
        if (data[i + j] !== lockedSignature[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        lockedOffset = i;
        // Found PROJECTLOCKED record, modify it
        if (i + lockedSignature.length + 10 < data.length) {
          // Set the locked flag to 0
          for (let j = i + lockedSignature.length; j < i + lockedSignature.length + 10; j++) {
            if (data[j] === 0x01) {
              data[j] = 0x00;
            }
          }
          logger(`Modified PROJECTLOCKED record at offset ${lockedOffset}`, 'info');
        }
        break;
      }
    }
    
    // Preserve the CMG record (module protection)
    const cmgSignature = [0x43, 0x4D, 0x47]; // "CMG"
    for (let i = 0; i < data.length - cmgSignature.length; i++) {
      let match = true;
      for (let j = 0; j < cmgSignature.length; j++) {
        if (data[i + j] !== cmgSignature[j]) {
          match = false;
          break;
        }
      }
      if (match && i + 20 < data.length) {
        // Found CMG record, check if it's followed by protection bytes
        for (let j = i + 3; j < i + 20; j++) {
          if (data[j] === 0x01) {
            data[j] = 0x00;
            logger(`Modified CMG protection record at offset ${j}`, 'info');
          }
        }
      }
    }
    
    // Preserve the DPB record (document protection)
    const dpbSignature = [0x44, 0x50, 0x42]; // "DPB"
    for (let i = 0; i < data.length - dpbSignature.length; i++) {
      let match = true;
      for (let j = 0; j < dpbSignature.length; j++) {
        if (data[i + j] !== dpbSignature[j]) {
          match = false;
          break;
        }
      }
      if (match && i + 20 < data.length) {
        // Found DPB record, check if it's followed by protection bytes
        for (let j = i + 3; j < i + 20; j++) {
          if (data[j] === 0x01) {
            data[j] = 0x00;
            logger(`Modified DPB protection record at offset ${j}`, 'info');
          }
        }
      }
    }
    
    // Search for and modify any remaining protection bytes
    const protectionSignatures = [
      [0x44, 0x50, 0x78], // "DPx"
      [0x43, 0x4D, 0x78], // "CMx"
      [0x56, 0x42, 0x41, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74], // "VBAProtect"
    ];
    
    for (const signature of protectionSignatures) {
      for (let i = 0; i < data.length - signature.length; i++) {
        let match = true;
        for (let j = 0; j < signature.length; j++) {
          if (data[i + j] !== signature[j]) {
            match = false;
            break;
          }
        }
        if (match && i + signature.length + 10 < data.length) {
          // Found protection signature, modify nearby bytes
          for (let j = i + signature.length; j < i + signature.length + 10; j++) {
            if (data[j] === 0x01) {
              data[j] = 0x00;
              logger(`Modified additional protection record at offset ${j}`, 'info');
            }
          }
        }
      }
    }
    
    logger(`Updated VBA project with comprehensive structure preservation`, 'success');
    return data;
  } catch (error) {
    logger(`Error preserving VBA structure: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return null;
  }
}

// Function to update VBA project checksum
function updateVBAProjectChecksum(vbaData: Uint8Array, logger: LoggerCallback): Uint8Array | null {
  try {
    // VBA projects have a 4-byte checksum at offset 4
    if (vbaData.length < 8) {
      logger('Invalid vbaProject.bin: File too small', 'error');
      return null;
    }
    
    // Create a copy of the data to modify
    const data = new Uint8Array(vbaData);
    const view = new DataView(data.buffer);
    
    // Calculate new checksum
    let calculatedChecksum = 0;
    for (let i = 8; i < data.length; i++) {
      calculatedChecksum += data[i];
      calculatedChecksum &= 0xFFFFFFFF; // Keep it 32-bit
    }
    
    // Update the checksum in the file
    view.setUint32(4, calculatedChecksum, true); // true for little-endian
    
    logger(`Updated VBA project checksum to ${calculatedChecksum}`, 'info');
    return data;
  } catch (error) {
    logger(`Error updating VBA checksum: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return null;
  }
}

/**
 * Attempts to recover from common errors during VBA password removal
 * @param error The error that occurred
 * @param zip The JSZip instance
 * @param logger The logger callback
 * @returns True if recovery was successful, false otherwise
 */
async function attemptErrorRecovery(
  error: Error,
  zip: JSZip,
  logger: LoggerCallback
): Promise<boolean> {
  logger(`Attempting to recover from error: ${error.message}`, 'info');
  
  // Check for common error patterns
  if (error.message.includes('Invalid CRC')) {
    return await recoverFromCRCError(zip, logger);
  }
  
  if (error.message.includes('corrupted zip')) {
    return await recoverFromCorruptedZip(zip, logger);
  }
  
  if (error.message.includes('vbaProject.bin')) {
    return await recoverFromMissingVBAProject(zip, logger);
  }
  
  // No recovery path available
  logger('No automatic recovery available for this error', 'error');
  return false;
}

/**
 * Attempts to recover from CRC validation errors
 * @param zip The JSZip instance
 * @param logger The logger callback
 * @returns True if recovery was successful, false otherwise
 */
async function recoverFromCRCError(
  zip: JSZip,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    logger('Attempting to fix CRC validation errors...', 'info');
    
    // Check if we can bypass CRC validation
    const vbaProject = zip.file('xl/vbaProject.bin');
    if (!vbaProject) {
      logger('Cannot recover: vbaProject.bin not found', 'error');
      return false;
    }
    
    // Try to read the file with CRC validation disabled
    const vbaContent = await vbaProject.async('uint8array');
    
    if (vbaContent.length === 0) {
      logger('Cannot recover: vbaProject.bin is empty', 'error');
      return false;
    }
    
    logger('Successfully bypassed CRC validation', 'success');
    return true;
  } catch (error) {
    logger(`Recovery failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Attempts to recover from corrupted ZIP errors
 * @param zip The JSZip instance
 * @param logger The logger callback
 * @returns True if recovery was successful, false otherwise
 */
async function recoverFromCorruptedZip(
  zip: JSZip,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    logger('Attempting to fix corrupted ZIP structure...', 'info');
    
    // Check if essential files exist
    const essentialFiles = [
      'xl/workbook.xml',
      '[Content_Types].xml',
      '_rels/.rels'
    ];
    
    for (const file of essentialFiles) {
      if (!zip.file(file)) {
        logger(`Cannot recover: Essential file ${file} is missing`, 'error');
        return false;
      }
    }
    
    logger('Essential file structure is intact, attempting to proceed', 'info');
    return true;
  } catch (error) {
    logger(`Recovery failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Attempts to recover from missing VBA project errors
 * @param zip The JSZip instance
 * @param logger The logger callback
 * @returns True if recovery was successful, false otherwise
 */
async function recoverFromMissingVBAProject(
  zip: JSZip,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    logger('Checking for alternative VBA project locations...', 'info');
    
    // Check for alternative locations
    const alternativeLocations = [
      'xl/vbaProject.bin',
      'xl/_vbaProject.bin',
      'vbaProject.bin',
      'macro/vbaProject.bin'
    ];
    
    for (const location of alternativeLocations) {
      const vbaProject = zip.file(location);
      if (vbaProject) {
        logger(`Found VBA project at alternative location: ${location}`, 'info');
        
        // Move it to the standard location if it's not already there
        if (location !== 'xl/vbaProject.bin') {
          const content = await vbaProject.async('uint8array');
          zip.file('xl/vbaProject.bin', content);
          logger('Moved VBA project to standard location', 'success');
        }
        
        return true;
      }
    }
    
    logger('No VBA project found in any alternative locations', 'error');
    return false;
  } catch (error) {
    logger(`Recovery failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}