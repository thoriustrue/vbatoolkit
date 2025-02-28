/**
 * Represents a VBA code module
 */
export interface VBAModule {
  /** The name of the module */
  name: string;
  /** The type of the module */
  type: VBAModuleType;
  /** The VBA code content */
  code: string;
}

/**
 * Types of VBA modules
 */
export enum VBAModuleType {
  /** Standard VBA module */
  Standard = 'Standard Module',
  /** Class module */
  Class = 'Class Module',
  /** UserForm */
  Form = 'UserForm',
  /** Document module (e.g. ThisWorkbook, Sheet1) */
  Document = 'Document Module',
  /** Unknown module type */
  Unknown = 'Unknown'
} 