'use strict';

import JSZip from 'jszip';
import { LoggerCallback } from './types';

// ZIP signature constant
export const ZIP_SIGNATURE = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

/**
 * Validates if a buffer contains a valid ZIP file by checking its signature
 */
export function isValidZip(buffer: ArrayBuffer): boolean {
  const header = new Uint8Array(buffer.slice(0, 4));
  return arraysEqual(header, ZIP_SIGNATURE);
}

/**
 * Helper function to compare two Uint8Arrays
 */
function arraysEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  return a.every((val, i) => val === b[i]);
}

/**
 * Validates a ZIP file with additional logging
 */
export async function validateZipFile(
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    if (!isValidZip(fileData)) {
      logger('Invalid ZIP file format', 'error');
      return false;
    }
    
    logger('ZIP file validation passed', 'info');
    return true;
  } catch (error) {
    logger(`ZIP validation error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Validates Office file structure
 * @param zip The JSZip instance
 * @param logger Callback function for logging messages
 * @returns True if structure is valid, false otherwise
 */
export function validateOfficeStructure(zip: JSZip, logger: LoggerCallback): boolean {
  try {
    // Check for essential Office files
    const requiredPaths = [
      '[Content_Types].xml',
      '_rels/.rels'
    ];
    
    const missingPaths = requiredPaths.filter(path => !zip.files[path]);
    
    if (missingPaths.length > 0) {
      logger(`Missing required Office files: ${missingPaths.join(', ')}`, 'error');
      return false;
    }
    
    // Check for Excel-specific files
    const isExcel = zip.files['xl/workbook.xml'] !== undefined;
    if (isExcel) {
      logger('Valid Excel file structure detected', 'info');
    } else {
      logger('Not a valid Excel file structure', 'error');
      return false;
    }
    
    return true;
  } catch (error) {
    logger(`Structure validation error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Validates Office file structure and required components
 */
export function validateOfficeCRC(zip: JSZip, logger: LoggerCallback): boolean {
  try {
    // Check for required Office files
    const requiredFiles = [
      '[Content_Types].xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels'
    ];
    
    const missingFiles = requiredFiles.filter(f => !zip.files[f]);
    if (missingFiles.length > 0) {
      logger(`Missing required files: ${missingFiles.join(', ')}`, 'error');
      return false;
    }
    
    // Validate ZIP structure
    const vbaProject = zip.files['xl/vbaProject.bin'];
    if (!vbaProject) {
      logger('No VBA project found in file', 'error');
      return false;
    }
    
    return true;
  } catch (error) {
    logger(`CRC validation failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

export function validateExcelStructure(zip: JSZip, logger: LoggerCallback) {
  const requiredEntries = [
    '[Content_Types].xml',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/sheet1.xml'
  ];

  requiredEntries.forEach(entry => {
    if (!zip.files[entry]) {
      logger(`MISSING CRITICAL ENTRY: ${entry}`, 'error');
    }
  });

  // Validate workbook XML root element
  const workbookEntry = zip.files['xl/workbook.xml'];
  if (workbookEntry) {
    const content = workbookEntry.async('string');
    if (!content.includes('<workbook xmlns=')) {
      logger('Invalid workbook.xml: Missing root namespace declaration', 'error');
    }
  }
}

// Re-export the function
export const isValidZip = isValidZip;

export function validateOfficeStructure(zip: JSZip, logger: LoggerCallback): boolean {
  // Implementation of validateOfficeStructure
  // This function is now implemented in the validateOfficeStructure method
  return true; // Placeholder return, actual implementation needed
}

export function validateExcelStructure(zip: JSZip, logger: LoggerCallback) {
  // Implementation of validateExcelStructure
  // This function is now implemented in the validateExcelStructure method
}

export function validateOfficeStructure(zip: JSZip, logger: LoggerCallback): boolean {
  // Implementation of validateOfficeStructure
  // This function is now implemented in the validateOfficeStructure method
  return true; // Placeholder return, actual implementation needed
}

export function validateExcelStructure(zip: JSZip, logger: LoggerCallback) {
  // Implementation of validateExcelStructure
  // This function is now implemented in the validateExcelStructure method
} 