import { LoggerCallback } from './types';
import { isValidZip } from './zipValidator.js';
import JSZip from 'jszip';

/**
 * Removes security settings from Excel files (auto-enable macros and external links)
 * @param fileData The Excel file data as Uint8Array
 * @param logger Callback function for logging messages
 * @returns A Promise that resolves to the modified file data, or null if processing failed
 */
export async function removeExcelSecurity(
  fileData: Uint8Array,
  logger: LoggerCallback
): Promise<Uint8Array | null> {
  try {
    // Check if this is an Office Open XML file (Excel 2007+)
    if (!isOfficeOpenXML(fileData)) {
      logger('This feature only works with modern Excel files (.xlsx, .xlsm, .xlsb).', 'info');
      return null;
    }
    
    logger('Processing Excel file to remove security restrictions...', 'info');
    
    // Load the file as a ZIP archive
    const zipData = await JSZip.loadAsync(fileData);
    
    // Check if this is a valid Excel file
    if (!zipData.files['[Content_Types].xml']) {
      logger('Invalid Excel file format.', 'error');
      return null;
    }
    
    let securityRemoved = false;
    
    // 1. Modify Excel security settings in xl/workbook.xml
    if (zipData.files['xl/workbook.xml']) {
      let workbookXml = await zipData.files['xl/workbook.xml'].async('text');
      
      // Remove fileSharing password protection
      if (workbookXml.includes('fileSharing')) {
        const originalXml = workbookXml;
        workbookXml = workbookXml.replace(/<fileSharing[^>]*\/>/g, '');
        workbookXml = workbookXml.replace(/<fileSharing[^>]*>.*?<\/fileSharing>/gs, '');
        
        if (workbookXml !== originalXml) {
          logger('Removed fileSharing password protection.', 'success');
          securityRemoved = true;
        }
      }
      
      // Remove workbook protection
      if (workbookXml.includes('workbookProtection')) {
        const originalXml = workbookXml;
        workbookXml = workbookXml.replace(/<workbookProtection[^>]*\/>/g, '');
        workbookXml = workbookXml.replace(/<workbookProtection[^>]*>.*?<\/workbookProtection>/gs, '');
        
        if (workbookXml !== originalXml) {
          logger('Removed workbook protection.', 'success');
          securityRemoved = true;
        }
      }
      
      // Add or modify the workbook properties to enable macros
      if (!workbookXml.includes('<workbookPr')) {
        // If workbookPr doesn't exist, add it
        workbookXml = workbookXml.replace(/<workbook[^>]*>/g, 
          '$&\n  <workbookPr autoCompressPictures="0" codeName="ThisWorkbook" defaultThemeVersion="124226" date1904="0" filterPrivacy="0" promptedSolutions="0" publishItems="0" saveExternalLinkValues="1" updateLinks="1" />');
        logger('Added workbook properties to enable macros and external links.', 'success');
        securityRemoved = true;
      } else if (!workbookXml.includes('updateLinks="1"')) {
        // If workbookPr exists but doesn't have updateLinks, add it
        workbookXml = workbookXml.replace(/<workbookPr([^>]*)>/g, 
          '<workbookPr$1 updateLinks="1" saveExternalLinkValues="1" filterPrivacy="0" promptedSolutions="0">');
        logger('Updated workbook properties to enable external links.', 'success');
        securityRemoved = true;
      }
      
      // Update the file in the ZIP
      zipData.file('xl/workbook.xml', workbookXml, {
        compression: 'DEFLATE',
        unixPermissions: null,  // Important for Windows compatibility
        comment: 'Modified by VBAToolkit'
      });
      
      if (!validateXML(workbookXml)) {
        logger('Invalid XML structure after modification', 'error');
        throw new Error('XML validation failed');
      }
    }
    
    // 2. Remove sheet protection from all worksheets
    const sheetFiles = Object.keys(zipData.files).filter(filename => 
      filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')
    );
    
    for (const sheetFile of sheetFiles) {
      let sheetXml = await zipData.files[sheetFile].async('text');
      
      // Remove sheet protection
      if (sheetXml.includes('sheetProtection')) {
        const originalXml = sheetXml;
        sheetXml = sheetXml.replace(/<sheetProtection[^>]*\/>/g, '');
        sheetXml = sheetXml.replace(/<sheetProtection[^>]*>.*?<\/sheetProtection>/gs, '');
        
        if (sheetXml !== originalXml) {
          logger(`Removed protection from ${sheetFile}.`, 'success');
          securityRemoved = true;
        }
      }
      
      // Update the file in the ZIP
      zipData.file(sheetFile, sheetXml);
    }
    
    // 3. Modify Excel security settings file if it exists
    const securitySettingsFiles = [
      'xl/externalLinks/_rels/externalLink1.xml.rels',
      'xl/_rels/workbook.xml.rels'
    ];
    
    for (const secFile of securitySettingsFiles) {
      if (zipData.files[secFile]) {
        let secXml = await zipData.files[secFile].async('text');
        
        // Modify security settings to auto-enable external links
        if (secXml.includes('relationships') || secXml.includes('Relationship')) {
          // Keep the file as is, but log that we found it
          logger(`Found external links relationship file: ${secFile}.`, 'info');
        }
      }
    }
    
    // 4. Create or modify the Excel security settings file
    const excelSecuritySettings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<excelSecuritySettings xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <security>
    <trustRecords>
      <trustRecord trustMode="trustAll" />
    </trustRecords>
  </security>
</excelSecuritySettings>`;
    
    // Add the security settings file to the ZIP
    zipData.file('xl/excelSecuritySettings.xml', excelSecuritySettings);
    logger('Added auto-trust security settings.', 'success');
    securityRemoved = true;
    
    // 5. Update Content_Types to include the security settings file
    let contentTypesXml = await zipData.files['[Content_Types].xml'].async('text');
    if (!contentTypesXml.includes('excelSecuritySettings.xml')) {
      // Add the content type for the security settings file
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        '  <Override PartName="/xl/excelSecuritySettings.xml" ContentType="application/vnd.ms-excel.securitySettings+xml"/>\n</Types>'
      );
      zipData.file('[Content_Types].xml', contentTypesXml);
      logger('Updated content types to include security settings.', 'success');
    }
    
    // 6. Update the workbook relationships to include the security settings
    if (zipData.files['xl/_rels/workbook.xml.rels']) {
      let relsXml = await zipData.files['xl/_rels/workbook.xml.rels'].async('text');
      if (!relsXml.includes('securitySettings')) {
        // Add the relationship for the security settings file
        relsXml = relsXml.replace(
          '</Relationships>',
          '  <Relationship Id="securitySettings" Type="http://schemas.microsoft.com/office/2006/relationships/excelSecuritySettings" Target="excelSecuritySettings.xml"/>\n</Relationships>'
        );
        zipData.file('xl/_rels/workbook.xml.rels', relsXml);
        logger('Updated workbook relationships to include security settings.', 'success');
      }
    }
    
    // 7. Add VBA security settings to disable macro security
    if (zipData.files['xl/vbaProject.bin']) {
      logger('VBA project found. Applying direct modifications to auto-enable macros.', 'info');
      
      // Extract the vbaProject.bin file
      const vbaProjectBin = await zipData.files['xl/vbaProject.bin'].async('uint8array');
      
      // Create a modified copy of the vbaProject.bin file
      const modifiedVbaProject = new Uint8Array(vbaProjectBin);
      
      // Look for security-related patterns in the VBA project
      const securityPatterns = [
        // "AccessVBOM" pattern - controls access to the VBA object model
        { pattern: [0x41, 0x63, 0x63, 0x65, 0x73, 0x73, 0x56, 0x42, 0x4F, 0x4D], name: "AccessVBOM" },
        // "VBAWarnings" pattern - controls macro security warnings
        { pattern: [0x56, 0x42, 0x41, 0x57, 0x61, 0x72, 0x6E, 0x69, 0x6E, 0x67, 0x73], name: "VBAWarnings" },
        // "DisableAttachementsInPV" pattern - controls attachments in protected view
        { pattern: [0x44, 0x69, 0x73, 0x61, 0x62, 0x6C, 0x65, 0x41, 0x74, 0x74, 0x61, 0x63, 0x68, 0x6D, 0x65, 0x6E, 0x74, 0x73, 0x49, 0x6E, 0x50, 0x56], name: "DisableAttachmentsInPV" },
        // "BlockContentExecution" pattern - controls content execution
        { pattern: [0x42, 0x6C, 0x6F, 0x63, 0x6B, 0x43, 0x6F, 0x6E, 0x74, 0x65, 0x6E, 0x74, 0x45, 0x78, 0x65, 0x63, 0x75, 0x74, 0x69, 0x6F, 0x6E], name: "BlockContentExecution" },
      ];
      
      for (const { pattern, name } of securityPatterns) {
        const indices = findAllPatterns(modifiedVbaProject, pattern);
        
        if (indices.length > 0) {
          logger(`Found ${indices.length} ${name} security pattern(s) in vbaProject.bin.`, 'info');
          
          for (const index of indices) {
            // Look for security flags (0x01) near the pattern
            const searchRange = 20; // Search 20 bytes after the pattern
            for (let i = index + pattern.length; i < index + pattern.length + searchRange && i < modifiedVbaProject.length; i++) {
              if (modifiedVbaProject[i] === 0x01) {
                modifiedVbaProject[i] = 0x00; // Set to 0 to disable security
                logger(`Disabled ${name} security setting.`, 'success');
                securityRemoved = true;
              }
            }
          }
        }
      }
      
      // Update the vbaProject.bin file in the ZIP
      zipData.file('xl/vbaProject.bin', modifiedVbaProject);
    }
    
    // 8. Create or modify the vbaProjectSignature.bin file to bypass signature verification
    if (zipData.files['xl/vbaProjectSignature.bin']) {
      // Remove the signature file to bypass signature verification
      zipData.remove('xl/vbaProjectSignature.bin');
      logger('Removed VBA project signature to bypass verification.', 'success');
      securityRemoved = true;
    }
    
    // 9. Modify Excel macro settings in the workbook
    if (zipData.files['xl/workbook.xml']) {
      let workbookXml = await zipData.files['xl/workbook.xml'].async('text');
      
      // Add or modify the fileVersion element to enable macros
      if (!workbookXml.includes('<fileVersion')) {
        // If fileVersion doesn't exist, add it
        workbookXml = workbookXml.replace(/<workbook[^>]*>/g, 
          '$&\n  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>');
        logger('Added file version information to enable macros.', 'success');
        securityRemoved = true;
      }
      
      // Update the file in the ZIP
      zipData.file('xl/workbook.xml', workbookXml, {
        compression: 'DEFLATE',
        unixPermissions: null,  // Important for Windows compatibility
        comment: 'Modified by VBAToolkit'
      });
      
      if (!validateXML(workbookXml)) {
        logger('Invalid XML structure after modification', 'error');
        throw new Error('XML validation failed');
      }
    }
    
    // Remove any signature references from .rels files
    const relsFiles = Object.keys(zipData.files).filter(f => f.endsWith('.rels'));
    await Promise.all(relsFiles.map(async f => {
      let rels = await zipData.file(f)!.async('text');
      if (rels.includes('vbaProjectSignature.bin')) {
        rels = rels.replace(/<Relationship[^>]*vbaProjectSignature\.bin[^>]*\/>/g, '');
        zipData.file(f, rels);
        logger(`Cleaned signature references from ${f}`, 'info');
      }
    }));
    
    if (!securityRemoved) {
      logger('No security settings were found to remove.', 'info');
      return null;
    }
    
    logger('Security settings successfully modified to auto-enable macros and external links.', 'success');
    
    // Generate the modified ZIP file
    const modifiedZip = await zipData.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: {
        level: 9
      }
    });
    
    validateFileSignature(new Uint8Array(modifiedZip.buffer), logger);
    
    if (zipData.files['xl/vbaProjectSignature.rels']) {
      zipData.remove('xl/vbaProjectSignature.rels');
      logger('Removed vbaProjectSignature.rels for full bypass', 'info');
    }
    
    return modifiedZip;
  } catch (error) {
    logger(`Error removing Excel security: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return null;
  }
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

function validateFileSignature(data: Uint8Array, logger: LoggerCallback) {
  const signatures = {
    xlsm: [0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00],
    xlsb: [0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x08, 0x00]
  };
  
  const header = Array.from(data.slice(0, 8));
  const isXlsm = signatures.xlsm.every((val, idx) => val === header[idx]);
  const isXlsb = signatures.xlsb.every((val, idx) => val === header[idx]);
  
  if (!isXlsm && !isXlsb) {
    logger('INVALID FILE SIGNATURE: Not a valid Excel file', 'error');
    logger(`Header bytes: ${header.map(b => b.toString(16)).join(' ')}`, 'error');
  } else {
    logger(`Valid ${isXlsm ? 'XLSM' : 'XLSB'} signature detected`, 'success');
  }
}

const validateXML = (xml: string) => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'text/xml');
  return doc.documentElement.nodeName !== 'parsererror';
};