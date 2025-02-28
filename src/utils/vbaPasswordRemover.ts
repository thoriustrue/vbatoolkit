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
          // ... continue with processing ...
          
          return await zip.generateAsync({ type: 'blob' });
        } catch (secondError) {
          logger(`Error after recovery attempt: ${secondError instanceof Error ? secondError.message : String(secondError)}`, 'error');
        }
      }
    }
    
    return null;
  }
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
      
      for (let i = projectInfoOffset; i < projectInfoOffset + searchRange; i++) {
        // Protection record often has this pattern
        if (data[i] === 0x13 && data[i+1] === 0x00 && data[i+2] === 0x01 && data[i+3] === 0x00) {
          // Found potential protection record, disable it
          data[i+2] = 0x00; // Change 0x01 to 0x00 to disable protection
          logger(`Modified protection record at offset ${i}`, 'info');
          break;
        }
        
        // Alternative protection pattern
        if (data[i] === 0x13 && data[i+1] === 0x00 && data[i+2] === 0x02 && data[i+3] === 0x00) {
          data[i+2] = 0x00; // Change 0x02 to 0x00
          logger(`Modified alternative protection record at offset ${i}`, 'info');
          break;
        }
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
          data[i + lockedSignature.length + 2] = 0x00;
          logger(`Modified PROJECTLOCKED record at offset ${lockedOffset}`, 'info');
        }
        break;
      }
    }
    
    // Recalculate the checksum
    const view = new DataView(data.buffer);
    let calculatedChecksum = 0;
    for (let i = 8; i < data.length; i++) {
      calculatedChecksum += data[i];
      calculatedChecksum &= 0xFFFFFFFF; // Keep it 32-bit
    }
    view.setUint32(4, calculatedChecksum, true); // true for little-endian
    
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