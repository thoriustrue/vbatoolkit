import AdmZip from 'adm-zip';
import { LoggerCallback } from './types';

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
    const zip = new AdmZip(Buffer.from(fileData));
    
    // Check if the ZIP has valid entries
    const entries = zip.getEntries();
    if (entries.length === 0) {
      logger('ZIP file contains no entries', 'error');
      return false;
    }
    
    logger(`ZIP file contains ${entries.length} entries`, 'info');
    
    // Validate CRC checksums
    if (!validateOfficeCRC(zip, logger)) {
      logger('ZIP file has CRC validation errors', 'error');
      return false;
    }
    
    return true;
  } catch (error) {
    logger(`ZIP validation error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

/**
 * Validates CRC checksums in the ZIP file
 * @param zip The AdmZip instance
 * @param logger Callback function for logging messages
 * @returns True if all CRCs are valid, false otherwise
 */
export function validateOfficeCRC(zip: AdmZip, logger: LoggerCallback): boolean {
  try {
    let valid = true;
    
    zip.getEntries().forEach(entry => {
      try {
        // Skip directories
        if (entry.isDirectory) return;
        
        const headerCRC = entry.header.crc;
        // AdmZip calculates CRC when reading the file
        const content = entry.getData();
        
        // Log successful validation
        logger(`Validated CRC for ${entry.entryName}`, 'info');
      } catch (error) {
        logger(`CRC check failed for ${entry.entryName}: ${error instanceof Error ? error.message : String(error)}`, 'error');
        valid = false;
      }
    });
    
    return valid;
  } catch (error) {
    logger(`CRC validation error: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return false;
  }
}

export function validateExcelStructure(zip: AdmZip, logger: LoggerCallback) {
  const requiredEntries = [
    '[Content_Types].xml',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/sheet1.xml'
  ];

  requiredEntries.forEach(entry => {
    if (!zip.getEntry(entry)) {
      logger(`MISSING CRITICAL ENTRY: ${entry}`, 'error');
    }
  });

  // Validate workbook XML root element
  const workbookEntry = zip.getEntry('xl/workbook.xml');
  if (workbookEntry) {
    const content = zip.readAsText(workbookEntry);
    if (!content.includes('<workbook xmlns=')) {
      logger('Invalid workbook.xml: Missing root namespace declaration', 'error');
    }
  }
} 