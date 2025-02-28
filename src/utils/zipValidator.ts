import AdmZip from 'adm-zip';
import { LoggerCallback } from './types';
import { signed } from 'crc-32';

export async function validateZipFile(data: ArrayBuffer, logger: LoggerCallback) {
  try {
    const zip = new AdmZip(Buffer.from(data));
    validateExcelStructure(zip, logger);
    const zipEntries = zip.getEntries();
    
    logger(`ZIP contains ${zipEntries.length} entries`, 'info');
    
    zipEntries.forEach(entry => {
      const entryInfo = [
        `Entry: ${entry.entryName}`,
        `Compressed: ${entry.header.sizeCompressed} bytes`,
        `Uncompressed: ${entry.header.size} bytes`,
        `Method: ${entry.header.method === 0 ? 'STORE' : 'DEFLATE'}`
      ].join(' | ');
      
      logger(entryInfo, 'info');
    });
    
    const comment = zip.getZipComment();
    logger(`ZIP comment: ${comment || 'None'}`, 'info');
    
  } catch (error) {
    logger(`ZIP validation failed: ${error.message}`, 'error');
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

export function validateOfficeCRC(zip: AdmZip, logger: LoggerCallback) {
  zip.getEntries().forEach(entry => {
    try {
      const content = zip.readFile(entry);
      const calculated = signed(Buffer.from(content)) >>> 0; // Convert to unsigned
      if (entry.header.crc !== calculated) {
        logger(`CRC mismatch in ${entry.entryName}: Expected 0x${entry.header.crc.toString(16)} vs Calculated 0x${calculated.toString(16)}`, 'error');
      }
    } catch (error) {
      logger(`CRC check failed for ${entry.entryName}: ${error.message}`, 'error');
    }
  });
} 