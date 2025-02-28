import JSZip from 'jszip';
import { LoggerCallback } from './types';

export async function injectVBACode(
  fileData: ArrayBuffer,
  logger: LoggerCallback
): Promise<void> {
  try {
    logger('Extracting Excel file...', 'info');
    const zip = await JSZip.loadAsync(fileData);

    // Log the contents of the ZIP archive
    logger(`ZIP contents: ${Object.keys(zip.files).join(', ')}`, 'info');

    // Extract vbaProject.bin
    const vbaProjectBin = await zip.file('xl/vbaProject.bin')?.async('uint8array');
    if (!vbaProjectBin) {
      logger('vbaProject.bin not found.', 'error');
      return;
    }

    logger('Modifying vbaProject.bin...', 'info');
    const modifiedVbaProject = modifyVbaBinary(vbaProjectBin);

    // Replace the modified vbaProject.bin back into the zip
    zip.file('xl/vbaProject.bin', modifiedVbaProject);

    // Generate the modified Excel file
    const modifiedFileData = await zip.generateAsync({ type: 'blob' });

    // Create download link using native browser API
    const url = URL.createObjectURL(modifiedFileData);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'modified_file.xlsm';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    logger('VBA code injected and file downloaded.', 'success');
  } catch (error) {
    logger(`Error modifying VBA project: ${error.message}`, 'error');
  }
}

function modifyVbaBinary(vbaProjectBin: Uint8Array): Uint8Array {
  // Look for password protection flags
  const passwordPatterns = [
    [0x44, 0x50, 0x42],  // DPB
    [0x43, 0x4D, 0x47],  // CMG
    [0x47, 0x5A, 0x49]   // GZ
  ];

  const modified = new Uint8Array(vbaProjectBin);
  
  passwordPatterns.forEach(pattern => {
    const indices = findAllPatterns(modified, pattern);
    indices.forEach(index => {
      // Nullify the password hash (16 bytes after pattern)
      for(let i = index + pattern.length; i < index + pattern.length + 16; i++) {
        if(i < modified.length) modified[i] = 0x00;
      }
    });
  });

  return modified;
} 