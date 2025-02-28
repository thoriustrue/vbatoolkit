import { Buffer } from 'buffer';
import * as XLSX from 'xlsx';

// Initialize Buffer for XLSX
globalThis.Buffer = Buffer;

export const readWorkbook = (data: ArrayBuffer, options?: XLSX.ParsingOptions) => {
  return XLSX.read(data, options);
};

export const writeWorkbook = (workbook: XLSX.WorkBook, options?: XLSX.WritingOptions) => {
  return XLSX.write(workbook, options);
}; 