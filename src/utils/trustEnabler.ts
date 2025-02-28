import JSZip from 'jszip';
import { LoggerCallback } from './types';

/**
 * Modifies Excel files to run with maximum trust settings by default
 */
export async function enableMaximumTrust(
  zip: JSZip,
  logger: LoggerCallback
): Promise<JSZip> {
  try {
    logger('Enabling maximum trust settings...', 'info');
    
    // 1. Add trusted document properties
    const corePropsFile = zip.file('docProps/core.xml');
    if (corePropsFile) {
      let content = await corePropsFile.async('string');
      
      // Add trusted document category
      if (!content.includes('<cp:category>')) {
        content = content.replace('</cp:coreProperties>',
          '  <cp:category>Trusted Document</cp:category>\n</cp:coreProperties>');
        zip.file('docProps/core.xml', content);
        logger('Added trusted document category', 'info');
      }
    }
    
    // 2. Modify workbook.xml to enable trust
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      let content = await workbookFile.async('string');
      
      // Add or modify workbookPr to enable trust
      if (content.includes('<workbookPr')) {
        content = content.replace(/<workbookPr([^>]*)>/g, 
          '<workbookPr$1 allowRefreshQuery="1" autoCompressPictures="0" defaultThemeVersion="124226" filterPrivacy="0" promptedSolutions="1" publishItems="1" saveExternalLinkValues="1" updateLinks="1">');
      } else {
        content = content.replace(/<workbook[^>]*>/g, 
          '$&\n  <workbookPr allowRefreshQuery="1" autoCompressPictures="0" defaultThemeVersion="124226" filterPrivacy="0" promptedSolutions="1" publishItems="1" saveExternalLinkValues="1" updateLinks="1"/>');
      }
      
      // Add fileVersion with trusted settings
      if (!content.includes('<fileVersion')) {
        content = content.replace(/<workbook[^>]*>/g, 
          '$&\n  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27166"/>');
      }
      
      zip.file('xl/workbook.xml', content);
      logger('Added trust settings to workbook.xml', 'info');
    }
    
    // 3. Add trusted VBA project settings
    const vbaProject = zip.file('xl/vbaProject.bin');
    if (vbaProject) {
      const vbaContent = await vbaProject.async('uint8array');
      
      // Look for the VBA project information section
      const projectInfoSignature = [0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74]; // "Project" in ASCII
      let projectInfoOffset = -1;
      
      for (let i = 0; i < vbaContent.length - projectInfoSignature.length; i++) {
        let match = true;
        for (let j = 0; j < projectInfoSignature.length; j++) {
          if (vbaContent[i + j] !== projectInfoSignature[j]) {
            match = false;
            break;
          }
        }
        if (match) {
          projectInfoOffset = i;
          break;
        }
      }
      
      if (projectInfoOffset !== -1) {
        // Create a copy of the VBA content
        const modifiedVba = new Uint8Array(vbaContent);
        
        // Set trusted flag in the project information
        // This is typically 100-200 bytes after the "Project" signature
        const searchRange = Math.min(300, modifiedVba.length - projectInfoOffset);
        
        for (let i = projectInfoOffset; i < projectInfoOffset + searchRange; i++) {
          // Look for security/trust flags
          if (modifiedVba[i] === 0x00 && modifiedVba[i+1] === 0x01 && 
              modifiedVba[i+2] === 0x01 && modifiedVba[i+3] === 0x00) {
            // Set to trusted (0x00, 0x00, 0x00, 0x00)
            modifiedVba[i+1] = 0x00;
            modifiedVba[i+2] = 0x00;
            logger('Set VBA project to trusted status', 'info');
            break;
          }
        }
        
        // Update the checksum
        let checksum = 0;
        for (let i = 8; i < modifiedVba.length; i++) {
          checksum += modifiedVba[i];
          checksum &= 0xFFFFFFFF;
        }
        
        const view = new DataView(modifiedVba.buffer);
        view.setUint32(4, checksum, true);
        
        zip.file('xl/vbaProject.bin', modifiedVba);
      }
    }
    
    // 4. Add trusted document settings in custom.xml
    let customPropsFile = zip.file('docProps/custom.xml');
    if (!customPropsFile) {
      // Create custom.xml if it doesn't exist
      const customXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" 
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_TrustedDocumentSettings">
    <vt:lpwstr>Trusted</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="_MarkAsTrusted">
    <vt:bool>true</vt:bool>
  </property>
</Properties>`;
      
      zip.file('docProps/custom.xml', customXml);
      
      // Update content types to include custom.xml
      const contentTypes = zip.file('[Content_Types].xml');
      if (contentTypes) {
        let content = await contentTypes.async('string');
        if (!content.includes('custom.xml')) {
          content = content.replace('</Types>', 
            '  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>\n</Types>');
          zip.file('[Content_Types].xml', content);
        }
      }
      
      // Update _rels/.rels to include custom.xml
      const rels = zip.file('_rels/.rels');
      if (rels) {
        let content = await rels.async('string');
        if (!content.includes('custom.xml')) {
          content = content.replace('</Relationships>', 
            '  <Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>\n</Relationships>');
          zip.file('_rels/.rels', content);
        }
      }
      
      logger('Added trusted document settings in custom.xml', 'info');
    }
    
    // 5. Remove any existing trust warnings
    const warningFiles = [
      'xl/vbaWarnings.xml',
      'xl/vbaProjectSignature.bin'
    ];
    
    for (const file of warningFiles) {
      if (zip.file(file)) {
        zip.remove(file);
        logger(`Removed trust warning file: ${file}`, 'info');
      }
    }
    
    logger('Maximum trust settings enabled', 'success');
    return zip;
  } catch (error) {
    logger(`Error enabling trust settings: ${error instanceof Error ? error.message : String(error)}`, 'error');
    return zip;
  }
} 