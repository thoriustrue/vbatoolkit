'use strict';

import JSZip from 'jszip';
import { LoggerCallback } from './types';
import { Buffer } from 'buffer';

/**
 * Validates a ZIP file structure
 * @param fileData The file data to validate
 * @param logger Callback function for logging messages
 * @returns True if validation passes, false otherwise
 */
export async function validateZipFile(
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<boolean> {
  try {
    // Use JSZip instead of AdmZip (browser-compatible)
    const zip = await JSZip.loadAsync(fileData);
    
    // Check if the ZIP has valid entries
    const entries = Object.keys(zip.files);
    if (entries.length === 0) {
      logger('ZIP file contains no entries', 'error');
      return false;
    }
    
    logger(`ZIP file contains ${entries.length} entries`, 'info');
    
    // Validate file structure
    if (!validateOfficeStructure(zip, logger)) {
      logger('ZIP file has invalid Office structure', 'error');
      return false;
    }
    
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
 * This function is replaced with validateOfficeStructure
 * Keeping the function signature for compatibility
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
    logger(`CRC validation failed: ${error.message}`, 'error');
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

// Replace Buffer with Uint8Array for browser compatibility
const ZIP_SIGNATURE = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

export function isValidZip(outputFileBuffer: ArrayBuffer): boolean {
  const headerBuffer = new Uint8Array(outputFileBuffer.slice(0, 4));
  return headerBuffer.length === ZIP_SIGNATURE.length &&
         headerBuffer.every((value, index) => value === ZIP_SIGNATURE[index]);
} 