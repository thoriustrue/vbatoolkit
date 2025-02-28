import AdmZip from 'adm-zip';
import { LoggerCallback } from './types';

export async function validateZipFile(data: ArrayBuffer, logger: LoggerCallback) {
  try {
    const zip = new AdmZip(Buffer.from(data));
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