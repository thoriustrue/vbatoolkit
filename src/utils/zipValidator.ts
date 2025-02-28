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
      logger('No VBA project found in file', 'warning');
      // Not returning false here as some Excel files might not have VBA
    }
    
    return true;
  } catch (error) {
    logger(`CRC validation failed: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Validates Office structure (alias for validateOfficeCRC for backward compatibility)
 */
export function validateOfficeStructure(zip: JSZip, logger: LoggerCallback): boolean {
  return validateOfficeCRC(zip, logger);
}

/**
 * Validates Excel-specific structure
 */
export function validateExcelStructure(zip: JSZip, logger: LoggerCallback): boolean {
  try {
    // Check for Excel-specific files
    if (!zip.files['xl/workbook.xml']) {
      logger('Missing workbook.xml - not a valid Excel file', 'error');
      return false;
    }
    
    // Check for worksheets
    const hasWorksheets = Object.keys(zip.files).some(
      filename => filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')
    );
    
    if (!hasWorksheets) {
      logger('No worksheets found in Excel file', 'error');
      return false;
    }
    
    logger('Excel structure validation passed', 'info');
    return true;
  } catch (error) {
    logger(`Excel validation error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
} 