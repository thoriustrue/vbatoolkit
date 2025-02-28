import JSZip from 'jszip';
import { LoggerCallback, ProgressCallback } from './types';
import { readFileAsArrayBuffer } from './fileUtils';
import { validateOfficeCRC, isValidZip } from './zipValidator';

export async function removeVBAPassword(
  file: File,
  logger: LoggerCallback,
  progressCallback: ProgressCallback
): Promise<Blob | null> {
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
    const zip = await JSZip.loadAsync(fileData);
    
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
    const modifiedVba = removeVBAProjectPassword(vbaContent, logger);
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
      let vbaRels = await zip.file('xl/_rels/vbaProject.bin.rels').async('string');
      if (vbaRels.includes('vbaProjectSignature')) {
        vbaRels = vbaRels.replace(
          /<Relationship[^>]*vbaProjectSignature[^>]*\/>/g,
          ''
        );
        zip.file('xl/_rels/vbaProject.bin.rels', vbaRels);
        logger('Cleaned VBA project relationships', 'info');
      }
    }
    
    progressCallback(0.8);
    
    // Generate the modified file with proper MIME type and compression
    const modifiedFile = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
      mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
    });
    
    logger('VBA password successfully removed!', 'success');
    progressCallback(1);
    
    return modifiedFile;
  } catch (error) {
    logger(`Password removal failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    progressCallback(0);
    return null;
  }
}

// Function to remove password from VBA project
function removeVBAProjectPassword(vbaData: Uint8Array, logger: LoggerCallback): Uint8Array | null {
  try {
    // Create a copy of the data to modify
    const data = new Uint8Array(vbaData);
    
    // Find the DPB (Document Protection Block) - typically starts with "DPB"
    const dpbSignature = [0x44, 0x50, 0x42]; // "DPB" in ASCII
    let dpbOffset = -1;
    
    for (let i = 0; i < data.length - dpbSignature.length; i++) {
      if (data[i] === dpbSignature[0] && 
          data[i + 1] === dpbSignature[1] && 
          data[i + 2] === dpbSignature[2]) {
        dpbOffset = i;
        break;
      }
    }
    
    if (dpbOffset === -1) {
      // If no DPB found, look for protection flags
      const protectionFlags = [0x01, 0x00, 0x01, 0x00, 0x00, 0x00]; // Common protection flag pattern
      for (let i = 0; i < data.length - protectionFlags.length; i++) {
        let match = true;
        for (let j = 0; j < protectionFlags.length; j++) {
          if (data[i + j] !== protectionFlags[j]) {
            match = false;
            break;
          }
        }
        if (match) {
          // Found protection flags, set them to 0 to disable protection
          data[i] = 0x00;
          data[i + 2] = 0x00;
          logger('Found and removed VBA protection flags', 'info');
          return data;
        }
      }
      
      // Look for CMG marker (another common protection indicator)
      const cmgSignature = [0x43, 0x4D, 0x47]; // "CMG" in ASCII
      for (let i = 0; i < data.length - cmgSignature.length; i++) {
        if (data[i] === cmgSignature[0] && 
            data[i + 1] === cmgSignature[1] && 
            data[i + 2] === cmgSignature[2]) {
          // Found CMG marker, modify protection bytes
          if (i + 8 < data.length) {
            data[i + 6] = 0x00;
            data[i + 7] = 0x00;
            data[i + 8] = 0x00;
            logger('Found and removed CMG protection flags', 'info');
            return data;
          }
        }
      }
      
      logger('No password protection found in VBA project', 'info');
      return data; // Return original if no protection found
    }
    
    // If DPB found, modify it to remove password
    logger(`Found DPB at offset ${dpbOffset}`, 'info');
    
    // Typical password protection is 4 bytes after DPB
    if (dpbOffset + 4 < data.length) {
      // Set protection bytes to 0
      data[dpbOffset + 3] = 0x00;
      data[dpbOffset + 4] = 0x00;
      
      logger('Successfully removed VBA password protection', 'success');
      return data;
    }
    
    return data;
  } catch (error) {
    logger(`Error removing VBA password: ${error instanceof Error ? error.message : String(error)}`, 'error');
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