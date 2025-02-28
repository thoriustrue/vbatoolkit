import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { removeExcelSecurity } from './excelSecurityRemover';

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
    logger(`Processing file: ${file.name} (${(file.size / 1024).toFixed(2)} KB)`, 'info');
    progressCallback(5);
    
    // Read the file as an ArrayBuffer
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const fileData = new Uint8Array(arrayBuffer);
    
    logger('File loaded successfully. Analyzing structure...', 'info');
    progressCallback(15);
    
    // First check if this is a valid Excel file
    if (!isValidExcelFile(fileData)) {
      logger('This does not appear to be a valid Excel file. Please ensure you are uploading an Excel file (.xlsm, .xls, .xlsb).', 'error');
      return null;
    }
    
    logger('Valid Excel file detected.', 'info');
    progressCallback(25);
    
    // Try to detect if the file has VBA content
    const hasVBA = detectVBAContent(fileData);
    if (!hasVBA) {
      logger('Warning: No clear VBA content detected in this file. Will still attempt to process it.', 'info');
    } else {
      logger('VBA content detected in the file.', 'info');
    }
    
    progressCallback(35);
    
    // For OOXML files, try the direct VBA project.bin approach first
    let modifiedData = fileData;
    let passwordRemoved = false;
    
    if (isOfficeOpenXML(fileData)) {
      logger('Attempting direct VBA project modification for Office Open XML file...', 'info');
      const directResult = await processVBAProjectBin(fileData, logger);
      
      if (directResult) {
        modifiedData = directResult;
        passwordRemoved = true;
        logger('Successfully modified VBA project using direct binary approach.', 'success');
      } else {
        logger('Direct VBA project modification failed, falling back to standard approach.', 'info');
      }
    }
    
    // If direct approach didn't work or it's not an OOXML file, use the standard approach
    if (!passwordRemoved) {
      // Use a more careful approach to preserve file integrity
      const standardResult = await safePasswordRemoval(modifiedData, logger, progressCallback);
      
      if (standardResult) {
        modifiedData = standardResult;
        logger('Password protection removed using standard approach.', 'success');
      } else {
        logger('Failed to process the file. No password signatures found or unsupported format.', 'error');
        return null;
      }
    }
    
    // Apply Excel security removal to auto-enable macros and external links
    logger('Applying security settings to auto-enable macros and external links...', 'info');
    progressCallback(85);
    
    // Try to remove Excel security settings
    const securityRemovedData = await removeExcelSecurity(modifiedData, logger);
    
    // Use the security-removed data if available, otherwise use the password-removed data
    const finalData = securityRemovedData || modifiedData;
    
    logger('File processing completed. Creating output file...', 'success');
    progressCallback(95);
    
    // Return the modified file with the original file type
    return new Blob([finalData], { type: file.type });
  } catch (error) {
    logger(`Error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return null;
  }
}

/**
 * Process the vbaProject.bin file directly to remove password protection
 * @param fileData The Excel file data
 * @param logger Callback function for logging messages
 * @returns The modified file data, or null if processing failed
 */
async function processVBAProjectBin(
  fileData: Uint8Array,
  logger: LoggerCallback
): Promise<Uint8Array | null> {
  try {
    // Load the file as a ZIP archive
    const zip = new JSZip();
    const zipData = await zip.loadAsync(fileData);
    
    // Check if the vbaProject.bin file exists
    if (!zipData.files['xl/vbaProject.bin']) {
      logger('No vbaProject.bin file found in the Excel file.', 'info');
      return null;
    }
    
    // Extract the vbaProject.bin file
    const vbaProjectBin = await zipData.files['xl/vbaProject.bin'].async('uint8array');
    logger('Found vbaProject.bin file. Size: ' + vbaProjectBin.length + ' bytes.', 'info');
    
    // Create a modified copy of the vbaProject.bin file
    const modifiedVbaProject = new Uint8Array(vbaProjectBin);
    let passwordRemoved = false;
    
    // Search for known password protection patterns in the vbaProject.bin file
    const protectionPatterns = [
      // DPB pattern (most common)
      { pattern: [0x44, 0x50, 0x42, 0x3D], name: "DPB=" },
      // CMG pattern
      { pattern: [0x43, 0x4D, 0x47, 0x3D], name: "CMG=" },
      // GC pattern
      { pattern: [0x47, 0x43, 0x3D], name: "GC=" },
      // DPx pattern with flag
      { pattern: [0x44, 0x50, 0x78, 0x01], name: "DPx" },
      // DPb pattern with flag
      { pattern: [0x44, 0x50, 0x62, 0x01], name: "DPb" },
      // DPI pattern with flag
      { pattern: [0x44, 0x50, 0x49, 0x01], name: "DPI" },
      // Project protection
      { pattern: [0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], name: "ProjectProtection" },
      // Password protection
      { pattern: [0x50, 0x61, 0x73, 0x73, 0x77, 0x6F, 0x72, 0x64, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], name: "PasswordProtection" },
    ];
    
    const validatePatternPosition = (index: number, buffer: Uint8Array) => {
      if (index > buffer.length - 100) {
        logger(`Suspicious pattern position at offset ${index}`, 'warning');
        return false;
      }
      return true;
    };
    
    for (const { pattern, name } of protectionPatterns) {
      const indices = findAllPatterns(modifiedVbaProject, pattern);
      
      if (indices.length > 0) {
        logger(`Found ${indices.length} ${name} pattern(s) in vbaProject.bin.`, 'info');
        
        for (const index of indices) {
          if (!validatePatternPosition(index, modifiedVbaProject)) continue;
          
          if (name.endsWith("=")) {
            // For patterns like DPB=, CMG=, GC=, clear the hash value
            let endIndex = index + pattern.length;
            while (endIndex < modifiedVbaProject.length && 
                   ((modifiedVbaProject[endIndex] >= 32 && modifiedVbaProject[endIndex] <= 126) || 
                    modifiedVbaProject[endIndex] === 0x0D || modifiedVbaProject[endIndex] === 0x0A)) {
              modifiedVbaProject[endIndex] = 0x20; // Replace with space
              endIndex++;
            }
          } else if (pattern[pattern.length - 1] === 0x01) {
            // For patterns ending with 0x01 flag, clear the flag
            modifiedVbaProject[index + pattern.length - 1] = 0x00;
            
            // Also clear a few bytes after the pattern to be safe
            for (let i = 0; i < 8; i++) {
              if (index + pattern.length + i < modifiedVbaProject.length) {
                modifiedVbaProject[index + pattern.length + i] = 0x00;
              }
            }
          } else {
            // For other patterns, look for 0x01 flags nearby
            const searchRange = 20; // Search 20 bytes after the pattern
            for (let i = index + pattern.length; i < index + pattern.length + searchRange && i < modifiedVbaProject.length; i++) {
              if (modifiedVbaProject[i] === 0x01) {
                modifiedVbaProject[i] = 0x00;
              }
            }
          }
          
          passwordRemoved = true;
        }
      }
    }
    
    // Additional specific patterns for VBA project protection
    const specificPatterns = [
      // Excel 2010-2013 protection pattern
      { pattern: [0x44, 0x50, 0x42, 0x3D, 0x22], name: "DPB=\"" },
      // Excel 2016-2019 protection pattern
      { pattern: [0x44, 0x50, 0x78, 0x01, 0x00, 0x00, 0x00], name: "DPx\\x01\\x00\\x00\\x00" },
      // Excel 365 protection pattern
      { pattern: [0x44, 0x50, 0x49, 0x01, 0x00, 0x00, 0x00], name: "DPI\\x01\\x00\\x00\\x00" },
      // Protection flag pattern
      { pattern: [0x56, 0x42, 0x41, 0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74, 0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], name: "VBAProjectProtection" },
    ];
    
    for (const { pattern, name } of specificPatterns) {
      const indices = findAllPatterns(modifiedVbaProject, pattern);
      
      if (indices.length > 0) {
        logger(`Found ${indices.length} ${name} specific pattern(s) in vbaProject.bin.`, 'info');
        
        for (const index of indices) {
          if (!validatePatternPosition(index, modifiedVbaProject)) continue;
          
          // For specific patterns, clear a larger area to ensure all protection is removed
          const clearRange = 50; // Clear 50 bytes after the pattern
          for (let i = index + pattern.length; i < index + pattern.length + clearRange && i < modifiedVbaProject.length; i++) {
            modifiedVbaProject[i] = 0x00;
          }
          
          passwordRemoved = true;
        }
      }
    }
    
    // Look for protection flags (0x01) near specific markers
    const protectionMarkers = [
      { pattern: [0x44, 0x50, 0x78], name: "DPx" }, // DPx
      { pattern: [0x44, 0x50, 0x62], name: "DPb" }, // DPb
      { pattern: [0x44, 0x50, 0x49], name: "DPI" }, // DPI
      { pattern: [0x50, 0x72, 0x6F, 0x74, 0x65, 0x63, 0x74, 0x69, 0x6F, 0x6E], name: "Protection" } // Protection
    ];
    
    for (const { pattern, name } of protectionMarkers) {
      const indices = findAllPatterns(modifiedVbaProject, pattern);
      
      if (indices.length > 0) {
        logger(`Found ${indices.length} ${name} marker(s) in vbaProject.bin.`, 'info');
        
        for (const index of indices) {
          if (!validatePatternPosition(index, modifiedVbaProject)) continue;
          
          // Check if there's a protection flag (0x01) right after the marker
          if (index + pattern.length < modifiedVbaProject.length) {
            // Clear several bytes after the marker to be safe
            for (let i = 0; i < 10; i++) {
              if (index + pattern.length + i < modifiedVbaProject.length) {
                if (modifiedVbaProject[index + pattern.length + i] === 0x01) {
                  modifiedVbaProject[index + pattern.length + i] = 0x00;
                  passwordRemoved = true;
                  logger(`Cleared protection flag after ${name} marker.`, 'success');
                }
              }
            }
          }
        }
      }
    }
    
    // If no password protection was found or removed, return null
    if (!passwordRemoved) {
      logger('No password protection patterns found in vbaProject.bin.', 'info');
      return null;
    }
    
    // Update the vbaProject.bin file in the ZIP
    zipData.file('xl/vbaProject.bin', modifiedVbaProject);
    
    // Generate the modified ZIP file
    const modifiedZip = await zipData.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: {
        level: 9
      }
    });
    
    logger('Successfully modified vbaProject.bin to remove password protection.', 'success');
    return modifiedZip;
  } catch (error) {
    logger(`Error processing vbaProject.bin: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return null;
  }
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