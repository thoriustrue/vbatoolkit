import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { removeExcelSecurity } from './excelSecurityRemover';
import { OfficeCrypto } from 'office-crypto';

// Type for the logger callback function
type LoggerCallback = (message: string, type: 'info' | 'error' | 'success') => void;
type ProgressCallback = (progress: number) => void;

/**
 * Removes VBA password protection from an Excel file
 * @param file The Excel file to process
 * @param logger Callback function for logging messages
 * @param progressCallback Callback function for reporting progress
 * @returns A Promise that resolves to the processed file as a Blob, or null if processing failed
 */
export async function removeVBAPassword(
  file: File,
  logger: LoggerCallback,
  progressCallback: ProgressCallback
): Promise<Blob | null> {
  try {
    const fileData = new Uint8Array(await readFileAsArrayBuffer(file));
    const zip = await JSZip.loadAsync(fileData);
    
    // Validate before processing
    const admZip = new AdmZip(Buffer.from(fileData));
    if (!validateOfficeCRC(admZip, logger)) {
      throw new Error('Invalid Office file structure');
    }

    // Process VBA container
    const processed = await processVBAContainer(fileData, logger);
    
    // Final validation
    const finalZip = new AdmZip(Buffer.from(processed));
    if (!validateOfficeCRC(finalZip, logger)) {
      throw new Error('Final file failed CRC validation');
    }

    return new Blob([processed], { 
      type: 'application/vnd.ms-excel.sheet.macroEnabled.12' 
    });
    
  } catch (error) {
    logger(`Password removal failed: ${error.message}`, 'error');
    return null;
  }
}

/**
 * Process the vbaProject.bin file directly to remove password protection
 * @param fileData The Excel file data
 * @param logger Callback function for logging messages
 * @returns The modified file data, or null if processing failed
 */
async function processVBAContainer(
  fileData: Uint8Array,
  logger: LoggerCallback
): Promise<Uint8Array> {
  const originalZip = await JSZip.loadAsync(fileData);
  
  // Preserve all original files and their compression settings
  const newZip = new JSZip();
  
  // 1. Copy all original files with their metadata
  originalZip.forEach((path, entry) => {
    newZip.file(path, entry, {
      compression: entry.options.compression,
      compressionOptions: entry.options.compressionOptions,
      date: entry.options.date,
      unixPermissions: entry.options.unixPermissions,
      dosPermissions: entry.options.dosPermissions
    });
  });

  // 2. Only modify vbaProject.bin
  const vbaProject = originalZip.file('xl/vbaProject.bin');
  if (!vbaProject) {
    logger('VBA project not found', 'error');
    throw new Error('Missing vbaProject.bin');
  }

  // 3. Process VBA project while preserving compression
  const originalCompression = vbaProject.options.compression;
  const decrypted = await OfficeCrypto.decrypt(
    await vbaProject.async('nodebuffer'),
    { type: 'agile' }
  );
  
  const processed = await OfficeCrypto.removeProtection(decrypted);
  const encrypted = await OfficeCrypto.encrypt(processed, { type: 'agile' });
  
  // 4. Update with original compression
  newZip.file('xl/vbaProject.bin', encrypted, {
    compression: originalCompression,
    date: vbaProject.options.date
  });

  // 5. Generate with original settings
  return newZip.generateAsync({
    type: 'uint8array',
    compression: 'STORE', // Force container-level store
    platform: 'DOS', // Required for Excel compatibility
    comment: originalZip.comment
  });
}

/**
 * Checks if the file is a valid Excel file
 * @param data The file data
 * @returns True if the file is a valid Excel file, false otherwise
 */
function isValidExcelFile(data: Uint8Array): boolean {
  // Check for OLE Compound File (Excel 97-2003, .xls)
  if (isOLECompoundFile(data)) {
    return true;
  }
  
  // Check for Office Open XML (Excel 2007+, .xlsx, .xlsm)
  if (isOfficeOpenXML(data)) {
    return true;
  }
  
  // Check for Excel Binary Workbook (.xlsb)
  if (isExcelBinary(data)) {
    return true;
  }
  
  return false;
}

/**
 * Checks if the file is an OLE Compound File (typical for Excel 97-2003)
 * @param data The file data
 * @returns True if the file is an OLE Compound File, false otherwise
 */
function isOLECompoundFile(data: Uint8Array): boolean {
  // OLE Compound File signature: D0 CF 11 E0 A1 B1 1A E1
  const signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  
  if (data.length < signature.length) {
    return false;
  }
  
  for (let i = 0; i < signature.length; i++) {
    if (data[i] !== signature[i]) {
      return false;
    }
  }
  
  return true;
}

/**
 * Checks if the file is an Office Open XML file (Excel 2007+)
 * @param data The file data
 * @returns True if the file is an Office Open XML file, false otherwise
 */
function isOfficeOpenXML(data: Uint8Array): boolean {
  // Office Open XML files are ZIP files, which start with PK signature
  // PK signature: 50 4B 03 04
  const signature = [0x50, 0x4B, 0x03, 0x04];
  
  if (data.length < signature.length) {
    return false;
  }
  
  for (let i = 0; i < signature.length; i++) {
    if (data[i] !== signature[i]) {
      return false;
    }
  }
  
  return true;
}

/**
 * Checks if the file is an Excel Binary Workbook (.xlsb)
 * @param data The file data
 * @returns True if the file is an Excel Binary Workbook, false otherwise
 */
function isExcelBinary(data: Uint8Array): boolean {
  // Excel Binary Workbook files are also ZIP files with PK signature
  // But we'll also look for specific XLSB content markers
  if (!isOfficeOpenXML(data)) {
    return false;
  }
  
  // Look for "workbook.bin" string which is common in XLSB files
  const workbookBinPattern = [0x77, 0x6F, 0x72, 0x6B, 0x62, 0x6F, 0x6F, 0x6B, 0x2E, 0x62, 0x69, 0x6E]; // "workbook.bin"
  return findPattern(data, workbookBinPattern) !== -1;
}

/**
 * Attempts to detect if the file contains VBA content
 * @param data The file data
 * @returns True if VBA content is detected, false otherwise
 */
function detectVBAContent(data: Uint8Array): boolean {
  // Common VBA-related strings to look for
  const vbaPatterns = [
    [0x56, 0x42, 0x41, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74], // "VBAProject"
    [0x5F, 0x56, 0x42, 0x41, 0x5F, 0x50, 0x52, 0x4F, 0x4A, 0x45, 0x43, 0x54], // "_VBA_PROJECT"
    [0x76, 0x62, 0x61, 0x2F], // "vba/"
    [0x56, 0x69, 0x73, 0x75, 0x61, 0x6C, 0x20, 0x42, 0x61, 0x73, 0x69, 0x63], // "Visual Basic"
  ];
  
  for (const pattern of vbaPatterns) {
    if (findPattern(data, pattern) !== -1) {
      return true;
    }
  }
  
  return false;
}

/**
 * A safer approach to remove password protection while preserving file integrity
 * @param fileData The original file data
 * @param logger Callback function for logging messages
 * @param progressCallback Callback function for reporting progress
 * @returns The modified file data, or null if processing failed
 */
async function safePasswordRemoval(
  fileData: Uint8Array,
  logger: LoggerCallback,
  progressCallback: ProgressCallback
): Promise<Uint8Array | null> {
  try {
    logger('Starting password removal process...', 'info');
    
    // Log initial file data size
    logger(`Initial file data size: ${fileData.length} bytes`, 'info');
    
    // Create a copy of the file data to modify
    const modifiedData = new Uint8Array(fileData);
    
    logger('Searching for password protection signatures...', 'info');
    progressCallback(40);
    
    // Track if we found and removed any password protection
    let passwordRemoved = false;
    
    // Determine file type for specialized handling
    let fileType: 'ole' | 'ooxml' | 'xlsb' | 'unknown' = 'unknown';
    
    if (isOLECompoundFile(fileData)) {
      fileType = 'ole';
      logger('Detected OLE Compound File format (.xls).', 'info');
    } else if (isOfficeOpenXML(fileData)) {
      if (isExcelBinary(fileData)) {
        fileType = 'xlsb';
        logger('Detected Excel Binary Workbook format (.xlsb).', 'info');
      } else {
        fileType = 'ooxml';
        logger('Detected Office Open XML format (.xlsm).', 'info');
      }
    }
    
    // PART 1: VBA Project Password Protection - Targeted approach based on file type
    
    if (fileType === 'ole') {
      // For OLE files (.xls), focus on the VBA project stream
      logger('Using OLE-specific approach for .xls files...', 'info');
      
      // Method 1: Search for "DPB=" signature (Document Protection Block)
      const dpbIndices = findAllPatterns(modifiedData, [0x44, 0x50, 0x42, 0x3D]); // "DPB="
      if (dpbIndices.length > 0) {
        logger(`Found ${dpbIndices.length} DPB signature(s).`, 'info');
        
        for (const dpbIndex of dpbIndices) {
          if (!validatePatternPosition(dpbIndex, modifiedData)) continue;
          
          // For OLE files, we need to be more careful with the modification
          // Only modify the bytes that are definitely part of the password hash
          // Typically, the password hash follows the DPB= signature
          
          // Find the end of the password hash (usually terminated by a null byte or a non-printable character)
          let endIndex = dpbIndex + 4;
          while (endIndex < modifiedData.length && 
                 ((modifiedData[endIndex] >= 32 && modifiedData[endIndex] <= 126) || // printable ASCII
                  modifiedData[endIndex] === 0x0D || modifiedData[endIndex] === 0x0A)) {
            endIndex++;
          }
          
          // Replace only the password hash part with zeros
          for (let i = dpbIndex + 4; i < endIndex; i++) {
            modifiedData[i] = 0;
          }
          passwordRemoved = true;
        }
        
        logger('Removed DPB password protection.', 'success');
      }
      
      // Method 2: Search for "CMG=" signature (newer Excel versions)
      const cmgIndices = findAllPatterns(modifiedData, [0x43, 0x4D, 0x47, 0x3D]); // "CMG="
      if (cmgIndices.length > 0) {
        logger(`Found ${cmgIndices.length} CMG signature(s).`, 'info');
        
        for (const cmgIndex of cmgIndices) {
          if (!validatePatternPosition(cmgIndex, modifiedData)) continue;
          
          // Find the end of the password hash
          let endIndex = cmgIndex + 4;
          while (endIndex < modifiedData.length && 
                 ((modifiedData[endIndex] >= 32 && modifiedData[endIndex] <= 126) || 
                  modifiedData[endIndex] === 0x0D || modifiedData[endIndex] === 0x0A)) {
            endIndex++;
          }
          
          // Replace only the password hash part with zeros
          for (let i = cmgIndex + 4; i < endIndex; i++) {
            modifiedData[i] = 0;
          }
          passwordRemoved = true;
        }
        
        logger('Removed CMG password protection.', 'success');
      }
      
      // Method 3: Search for protection flags near "VBAProject"
      const vbaProjectIndices = findAllPatterns(modifiedData, [0x56, 0x42, 0x41, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74]); // "VBAProject"
      if (vbaProjectIndices.length > 0) {
        logger(`Found ${vbaProjectIndices.length} VBAProject signature(s).`, 'info');
        
        for (const vbaIndex of vbaProjectIndices) {
          if (!validatePatternPosition(vbaIndex, modifiedData)) continue;
          
          // Common offsets for protection flags in OLE files
          const commonOffsets = [0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x4B, 0x4C, 0x4D, 0x4E, 0x4F, 
                                0x50, 0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x5B, 0x5C, 0x5D, 0x5E, 0x5F];
          
          for (const offset of commonOffsets) {
            if (vbaIndex + offset < modifiedData.length) {
              const value = modifiedData[vbaIndex + offset];
              if (value === 0x01 || value === 0xFF) {
                modifiedData[vbaIndex + offset] = 0;
                passwordRemoved = true;
                logger(`Cleared protection flag at offset +0x${offset.toString(16).toUpperCase()} from VBAProject.`, 'success');
              }
            }
          }
        }
      }
    } else if (fileType === 'ooxml' || fileType === 'xlsb') {
      logger('Using OOXML/XLSB-specific approach...', 'info');
      
      // Log the presence of vbaProject.bin
      const vbaProjectBinPattern = [0x76, 0x62, 0x61, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74, 0x2E, 0x62, 0x69, 0x6E];
      const vbaProjectBinIndices = findAllPatterns(modifiedData, vbaProjectBinPattern);
      
      logger(`Found ${vbaProjectBinIndices.length} vbaProject.bin reference(s).`, 'info');
      
      // Log each modification attempt
      for (const vbaIndex of vbaProjectBinIndices) {
        logger(`Processing vbaProject.bin at index: ${vbaIndex}`, 'info');
        
        // Search in a large window around the vbaProject.bin reference
        const searchWindow = 10000; // 10KB search window
        const startSearch = Math.max(0, vbaIndex - searchWindow);
        const endSearch = Math.min(modifiedData.length, vbaIndex + searchWindow);
        
        // Look for DPB= and CMG= in the search window
        for (let i = startSearch; i < endSearch - 4; i++) {
          // Check for "DPB="
          if (modifiedData[i] === 0x44 && modifiedData[i+1] === 0x50 && modifiedData[i+2] === 0x42 && modifiedData[i+3] === 0x3D) {
            logger(`Found DPB= signature near vbaProject.bin at offset ${i - vbaIndex} from reference.`, 'info');
            
            // Find the end of the password hash
            let endIndex = i + 4;
            while (endIndex < endSearch && 
                   ((modifiedData[endIndex] >= 32 && modifiedData[endIndex] <= 126) || 
                    modifiedData[endIndex] === 0x0D || modifiedData[endIndex] === 0x0A)) {
              endIndex++;
            }
            
            // Replace only the password hash part with zeros
            for (let j = i + 4; j < endIndex; j++) {
              modifiedData[j] = 0;
            }
            passwordRemoved = true;
          }
          
          // Check for "CMG="
          if (modifiedData[i] === 0x43 && modifiedData[i+1] === 0x4D && modifiedData[i+2] === 0x47 && modifiedData[i+3] === 0x3D) {
            logger(`Found CMG= signature near vbaProject.bin at offset ${i - vbaIndex} from reference.`, 'info');
            
            // Find the end of the password hash
            let endIndex = i + 4;
            while (endIndex < endSearch && 
                   ((modifiedData[endIndex] >= 32 && modifiedData[endIndex] <= 126) || 
                    modifiedData[endIndex] === 0x0D || modifiedData[endIndex] === 0x0A)) {
              endIndex++;
            }
            
            // Replace only the password hash part with zeros
            for (let j = i + 4; j < endIndex; j++) {
              modifiedData[j] = 0;
            }
            passwordRemoved = true;
          }
        }
        
        // Also look for protection flags (0x01 bytes) near specific markers
        const protectionMarkers = [
          [0x44, 0x50, 0x78], // DPx
          [0x44, 0x50, 0x62], // DPb
          [0x44, 0x50, 0x49], // DPI
          [0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E] // Protection
        ];
        
        for (const marker of protectionMarkers) {
          const markerIndices = findAllPatternsInRange(modifiedData, marker, startSearch, endSearch);
          
          for (const markerIndex of markerIndices) {
            // Check if there's a protection flag (0x01) right after the marker
            if (markerIndex + marker.length < endSearch && modifiedData[markerIndex + marker.length] === 0x01) {
              modifiedData[markerIndex + marker.length] = 0;
              passwordRemoved = true;
              logger(`Cleared protection flag after ${String.fromCharCode(...marker)} marker.`, 'success');
            }
          }
        }
        
        // Log the modified data size
        logger(`Modified data size: ${modifiedData.length} bytes`, 'info');
      }
    }
    
    progressCallback(60);
    
    // PART 2: Generic approach for all file types if no password was removed yet
    if (!passwordRemoved) {
      logger('No file-specific protection found. Trying generic approach...', 'info');
      
      // Method 1: Search for common protection signatures
      const protectionSignatures = [
        [0x44, 0x50, 0x42, 0x3D], // "DPB="
        [0x43, 0x4D, 0x47, 0x3D], // "CMG="
        [0x47, 0x43, 0x3D],       // "GC="
        [0x44, 0x50, 0x78, 0x01], // "DPx" with flag
        [0x44, 0x50, 0x62, 0x01], // "DPb" with flag
        [0x44, 0x50, 0x49, 0x01], // "DPI" with flag
      ];
      
      for (const signature of protectionSignatures) {
        const signatureIndices = findAllPatterns(modifiedData, signature);
        
        if (signatureIndices.length > 0) {
          logger(`Found ${signatureIndices.length} ${String.fromCharCode(...signature.filter(b => b >= 32 && b <= 126))} signature(s).`, 'info');
          
          for (const signatureIndex of signatureIndices) {
            if (signature[signature.length - 1] === 0x01) {
              // If the signature ends with a flag byte (0x01), just clear that byte
              modifiedData[signatureIndex + signature.length - 1] = 0;
            } else {
              // For other signatures (like DPB=), clear the following data (password hash)
              let endIndex = signatureIndex + signature.length;
              while (endIndex < modifiedData.length && 
                     ((modifiedData[endIndex] >= 32 && modifiedData[endIndex] <= 126) || 
                      modifiedData[endIndex] === 0x0D || modifiedData[endIndex] === 0x0A)) {
                endIndex++;
              }
              
              for (let i = signatureIndex + signature.length; i < endIndex; i++) {
                modifiedData[i] = 0;
              }
            }
            
            passwordRemoved = true;
          }
          
          logger(`Cleared ${String.fromCharCode(...signature.filter(b => b >= 32 && b <= 126))} protection.`, 'success');
        }
      }
      
      // Method 2: Search for "ProjectProtection" and "PasswordProtection" strings
      const protectionStrings = [
        [0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], // "ProjectProtection"
        [0x50, 0x61, 0x73, 0x73, 0x77, 0x6F, 0x72, 0x64, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], // "PasswordProtection"
      ];
      
      for (const protectionString of protectionStrings) {
        const stringIndices = findAllPatterns(modifiedData, protectionString);
        
        if (stringIndices.length > 0) {
          logger(`Found ${stringIndices.length} ${String.fromCharCode(...protectionString)} string(s).`, 'info');
          
          for (const stringIndex of stringIndices) {
            // Look for protection flags (0x01) in the vicinity
            const searchRange = 100; // Search 100 bytes after the string
            
            for (let i = stringIndex + protectionString.length; i < stringIndex + protectionString.length + searchRange && i < modifiedData.length; i++) {
              if (modifiedData[i] === 0x01) {
                modifiedData[i] = 0;
                passwordRemoved = true;
                logger(`Cleared protection flag near ${String.fromCharCode(...protectionString)}.`, 'success');
              }
            }
          }
        }
      }
    }
    
    progressCallback(80);
    
    // PART 3: Last resort approach - only if nothing else worked
    if (!passwordRemoved) {
      logger('No standard protection patterns found. Trying last resort approach...', 'info');
      
      // Look for VBA-related strings and check for protection flags nearby
      const vbaStrings = [
        [0x56, 0x42, 0x41], // "VBA"
        [0x56, 0x42, 0x41, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74], // "VBAProject"
        [0x76, 0x62, 0x61, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74], // "vbaProject"
      ];
      
      for (const vbaString of vbaStrings) {
        const stringIndices = findAllPatterns(modifiedData, vbaString);
        
        if (stringIndices.length > 0) {
          logger(`Found ${stringIndices.length} ${String.fromCharCode(...vbaString)} string(s).`, 'info');
          
          for (const stringIndex of stringIndices) {
            // Look for protection flags (0x01) in the vicinity
            const searchRange = 200; // Search 200 bytes around the string
            const startSearch = Math.max(0, stringIndex - searchRange);
            const endSearch = Math.min(modifiedData.length, stringIndex + vbaString.length + searchRange);
            
            for (let i = startSearch; i < endSearch; i++) {
              if (modifiedData[i] === 0x01) {
                // Check if this is likely a protection flag
                // (simple heuristic: check if surrounded by control bytes or zeros)
                let isLikelyFlag = true;
                for (let j = 1; j <= 3; j++) {
                  if (i - j >= startSearch && modifiedData[i - j] > 0x20 && modifiedData[i - j] < 0x7F) {
                    isLikelyFlag = false;
                    break;
                  }
                  if (i + j < endSearch && modifiedData[i + j] > 0x20 && modifiedData[i + j] < 0x7F) {
                    isLikelyFlag = false;
                    break;
                  }
                }
                
                if (isLikelyFlag) {
                  modifiedData[i] = 0;
                  passwordRemoved = true;
                  logger(`Cleared potential protection flag near ${String.fromCharCode(...vbaString)}.`, 'info');
                }
              }
            }
          }
        }
      }
    }
    
    // PART 4: Additional protection patterns for Excel 2010-2019
    // These are specific patterns found in newer Excel versions
    const modernProtectionPatterns = [
      // Excel 2010-2013 protection pattern
      [0x44, 0x50, 0x42, 0x3D, 0x22, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00], // "DPB="...
      // Excel 2016-2019 protection pattern
      [0x44, 0x50, 0x78, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x28, 0x00], // "DPx"...
      // Excel 365 protection pattern
      [0x44, 0x50, 0x49, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x3C, 0x00], // "DPI"...
    ];
    
    for (const pattern of modernProtectionPatterns) {
      const patternIndices = findAllPatterns(modifiedData, pattern);
      
      if (patternIndices.length > 0) {
        logger(`Found ${patternIndices.length} modern protection pattern(s).`, 'info');
        
        for (const patternIndex of patternIndices) {
          // Clear the protection flag (usually at offset 3 or 4)
          if (pattern[3] === 0x01) {
            modifiedData[patternIndex + 3] = 0;
          } else if (pattern[4] === 0x01) {
            modifiedData[patternIndex + 4] = 0;
          }
          
          // Also clear a few bytes after the pattern to be safe
          for (let i = 0; i < 4; i++) {
            if (patternIndex + pattern.length + i < modifiedData.length) {
              modifiedData[patternIndex + pattern.length + i] = 0;
            }
          }
          
          passwordRemoved = true;
        }
        
        logger('Cleared modern protection pattern(s).', 'success');
      }
    }
    
    // PART 5: Handle sheet protection in OOXML files
    if (fileType === 'ooxml' || fileType === 'xlsb') {
      // For OOXML files, we can also look for sheet protection markers
      const sheetProtectionPatterns = [
        [0x73, 0x68, 0x65, 0x65, 0x74, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], // "sheetProtection"
        [0x77, 0x6F, 0x72, 0x6B, 0x62, 0x6F, 0x6F, 0x6B, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], // "workbookProtection"
      ];
      
      for (const pattern of sheetProtectionPatterns) {
        const patternIndices = findAllPatterns(modifiedData, pattern);
        
        if (patternIndices.length > 0) {
          logger(`Found ${patternIndices.length} ${String.fromCharCode(...pattern)} marker(s).`, 'info');
          
          for (const patternIndex of patternIndices) {
            // Look for password hash attributes near the protection marker
            const searchRange = 200; // Search 200 bytes after the marker
            const endSearch = Math.min(modifiedData.length, patternIndex + pattern.length + searchRange);
            
            // Look for password="..." or algorithmName="..." attributes
            const passwordAttr = [0x70, 0x61, 0x73, 0x73, 0x77, 0x6F, 0x72, 0x64, 0x3D, 0x22]; // password="
            const algorithmAttr = [0x61, 0x6C, 0x67, 0x6F, 0x72, 0x69, 0x74, 0x68, 0x6D, 0x4E, 0x61, 0x6D, 0x65, 0x3D, 0x22]; // algorithmName="
            
            for (let i = patternIndex + pattern.length; i < endSearch - 10; i++) {
              // Check for password attribute
              let foundPasswordAttr = true;
              for (let j = 0; j < passwordAttr.length; j++) {
                if (i + j >= endSearch || modifiedData[i + j] !== passwordAttr[j]) {
                  foundPasswordAttr = false;
                  break;
                }
              }
              
              if (foundPasswordAttr) {
                // Find the closing quote
                let quoteEnd = i + passwordAttr.length;
                while (quoteEnd < endSearch && modifiedData[quoteEnd] !== 0x22) {
                  quoteEnd++;
                }
                
                // Replace the password hash with empty string
                for (let j = i + passwordAttr.length; j < quoteEnd; j++) {
                  modifiedData[j] = 0x20; // Replace with space
                }
                
                passwordRemoved = true;
                logger('Cleared sheet/workbook password hash.', 'success');
              }
              
              // Check for algorithm attribute
              let foundAlgorithmAttr = true;
              for (let j = 0; j < algorithmAttr.length; j++) {
                if (i + j >= endSearch || modifiedData[i + j] !== algorithmAttr[j]) {
                  foundAlgorithmAttr = false;
                  break;
                }
              }
              
              if (foundAlgorithmAttr) {
                // Find the closing quote
                let quoteEnd = i + algorithmAttr.length;
                while (quoteEnd < endSearch && modifiedData[quoteEnd] !== 0x22) {
                  quoteEnd++;
                }
                
                // Replace the algorithm name with empty string
                for (let j = i + algorithmAttr.length; j < quoteEnd; j++) {
                  modifiedData[j] = 0x20; // Replace with space
                }
                
                passwordRemoved = true;
                logger('Cleared sheet/workbook protection algorithm.', 'success');
              }
            }
          }
        }
      }
    }
    
    progressCallback(90);
    
    if (!passwordRemoved) {
      logger('No known password protection patterns found. The file might not be password-protected or uses an unsupported protection method.', 'info');
      // Even if no patterns were found, return the modified data anyway
      return modifiedData;
    }
    
    logger('Password protection removed successfully.', 'success');
    
    return modifiedData;
  } catch (error) {
    logger(`Error during password removal: ${error.message}`, 'error');
    return null;
  }
}

/**
 * Finds all occurrences of a byte pattern in a specific range of a Uint8Array
 * @param data The data to search in
 * @param pattern The pattern to search for
 * @param startIndex The start index of the range to search in
 * @param endIndex The end index of the range to search in
 * @returns An array of indices where the pattern was found
 */
function findAllPatternsInRange(data: Uint8Array, pattern: number[], startIndex: number, endIndex: number): number[] {
  const indices: number[] = [];
  
  for (let i = startIndex; i <= Math.min(endIndex, data.length) - pattern.length; i++) {
    let found = true;
    for (let j = 0; j < pattern.length; j++) {
      if (data[i + j] !== pattern[j]) {
        found = false;
        break;
      }
    }
    if (found) {
      indices.push(i);
    }
  }
  
  return indices;
}

/**
 * Reads a File as an ArrayBuffer
 * @param file The file to read
 * @returns A Promise that resolves to an ArrayBuffer
 */
function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      if (e.target?.result instanceof ArrayBuffer) {
        resolve(e.target.result);
      } else {
        reject(new Error('Failed to read file as ArrayBuffer'));
      }
    };
    reader.onerror = () => reject(new Error('File read error'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Finds all occurrences of a byte pattern in a Uint8Array
 * @param data The data to search in
 * @param pattern The pattern to search for
 * @returns An array of indices where the pattern was found
 */
function findAllPatterns(data: Uint8Array, pattern: number[]): number[] {
  const indices: number[] = [];
  
  for (let i = 0; i <= data.length - pattern.length; i++) {
    let found = true;
    for (let j = 0; j < pattern.length; j++) {
      if (data[i + j] !== pattern[j]) {
        found = false;
        break;
      }
    }
    if (found) {
      indices.push(i);
    }
  }
  
  return indices;
}

/**
 * Finds the first occurrence of a byte pattern in a Uint8Array
 * @param data The data to search in
 * @param pattern The pattern to search for
 * @returns The index of the first occurrence of the pattern, or -1 if not found
 */
function findPattern(data: Uint8Array, pattern: number[]): number {
  for (let i = 0; i <= data.length - pattern.length; i++) {
    let found = true;
    for (let j = 0; j < pattern.length; j++) {
      if (data[i + j] !== pattern[j]) {
        found = false;
        break;
      }
    }
    if (found) {
      return i;
    }
  }
  return -1;
}

function auditBinaryChanges(original: Uint8Array, modified: Uint8Array, logger: LoggerCallback) {
  const changes: string[] = [];
  
  for (let i = 0; i < Math.min(original.length, modified.length); i++) {
    if (original[i] !== modified[i]) {
      changes.push(
        `0x${i.toString(16).padStart(6, '0')}: ` +
        `0x${original[i].toString(16).padStart(2, '0')} → ` +
        `0x${modified[i].toString(16).padStart(2, '0')}`
      );
      
      if (changes.length > 50) break; // Limit output
    }
  }
  
  if (changes.length > 0) {
    logger(`Binary modifications detected (first ${changes.length}):\n${changes.join('\n')}`, 'info');
  } else {
    logger('No binary modifications made', 'warning');
  }
}