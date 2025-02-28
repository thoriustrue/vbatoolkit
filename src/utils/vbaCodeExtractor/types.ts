/**
 * Interface representing a VBA module
 */
export interface VBAModule {
  /** Name of the module */
  name: string;
  /** Type of the module */
  type: VBAModuleType;
  /** VBA code content */
  code: string;
  /** Whether the code extraction was successful */
  extractionSuccess: boolean;
}

/**
 * Enum representing different types of VBA modules
 */
export enum VBAModuleType {
  /** Standard module */
  Standard = 0,
  /** Class module */
  Class = 1,
  /** Form module */
  Form = 2,
  /** Document module (ThisWorkbook, Sheet1, etc.) */
  Document = 3,
  /** Unknown module type */
  Unknown = 4
} 