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
    
    progressCallback(0.6);
    
    // Replace the vbaProject.bin with the modified version
    zip.file('xl/vbaProject.bin', modifiedVba);
    
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
    
    progressCallback(0.8);
    
    // Generate the modified file
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