/**
 * Comprehensive test suite for VBA toolkit functionality
 * This can be used to validate that the enhancements work correctly
 */

import { removeVBAPassword } from '../utils/vbaPasswordRemover';
import { extractVBACode } from '../utils/vbaCodeExtractor/index';
import { validateVBAProject } from '../utils/vbaProtectionRemover';
import { extractVBACodeEnhanced } from '../utils/vbaCodeExtractor/enhancedExtractor';

interface TestResult {
  testName: string;
  success: boolean;
  message: string;
  duration: number;
}

export class VBAToolkitTester {
  private results: TestResult[] = [];
  private logger: (message: string, type: string) => void;

  constructor(logger: (message: string, type: string) => void) {
    this.logger = logger;
  }

  /**
   * Run a comprehensive test of VBA toolkit functionality
   */
  async runTests(): Promise<TestResult[]> {
    this.results = [];
    this.logger('Starting VBA Toolkit comprehensive tests...', 'info');

    // Test 1: VBA Binary Format Validation
    await this.testVBAValidation();

    // Test 2: Enhanced Protection Removal
    await this.testProtectionRemoval();

    // Test 3: Enhanced Code Extraction
    await this.testCodeExtraction();

    // Test 4: Error Handling
    await this.testErrorHandling();

    // Test 5: File Integrity
    await this.testFileIntegrity();

    this.logSummary();
    return this.results;
  }

  private async runTest(testName: string, testFn: () => Promise<void>): Promise<void> {
    const startTime = Date.now();
    try {
      await testFn();
      const duration = Date.now() - startTime;
      this.results.push({
        testName,
        success: true,
        message: 'Test passed',
        duration
      });
      this.logger(`✓ ${testName} (${duration}ms)`, 'success');
    } catch (error) {
      const duration = Date.now() - startTime;
      const message = error instanceof Error ? error.message : String(error);
      this.results.push({
        testName,
        success: false,
        message,
        duration
      });
      this.logger(`✗ ${testName}: ${message} (${duration}ms)`, 'error');
    }
  }

  private async testVBAValidation(): Promise<void> {
    await this.runTest('VBA Binary Format Validation', async () => {
      // Test with valid VBA binary data
      const validVBAData = new Uint8Array([
        // Mock VBA header
        0x01, 0x00, 0x00, 0x00, // signature
        0x10, 0x00, 0x00, 0x00, // checksum (placeholder)
        // Mock data that would sum to checksum
        0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x01, 0x00
      ]);

      // Calculate proper checksum
      let checksum = 0;
      for (let i = 8; i < validVBAData.length; i++) {
        checksum += validVBAData[i];
      }
      validVBAData[4] = checksum & 0xFF;
      validVBAData[5] = (checksum >> 8) & 0xFF;
      validVBAData[6] = (checksum >> 16) & 0xFF;
      validVBAData[7] = (checksum >> 24) & 0xFF;

      const isValid = validateVBAProject(validVBAData, this.logger);
      if (!isValid) {
        throw new Error('Valid VBA data was incorrectly marked as invalid');
      }

      // Test with invalid data
      const invalidVBAData = new Uint8Array([0x00, 0x00]);
      const isInvalid = validateVBAProject(invalidVBAData, this.logger);
      if (isInvalid) {
        throw new Error('Invalid VBA data was incorrectly marked as valid');
      }
    });
  }

  private async testProtectionRemoval(): Promise<void> {
    await this.runTest('Enhanced Protection Removal', async () => {
      // Create mock VBA data with protection record
      const protectedVBAData = new Uint8Array([
        // Header
        0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        // Mock project info
        0x50, 0x72, 0x6F, 0x6A, 0x65, 0x63, 0x74, // "Project"
        // Protection record
        0x13, 0x00, 0x01, 0x00, // This should be modified
        // Some additional data
        0x00, 0x00, 0x00, 0x00
      ]);

      // Import the function (this is a simplified test)
      // In a real test, we'd need to mock the actual VBA binary format
      
      // For now, just verify the function exists and doesn't crash
      const { removeVBAProtectionEnhanced } = await import('../utils/vbaProtectionRemover');
      const result = removeVBAProtectionEnhanced(protectedVBAData, this.logger);
      
      if (!result) {
        throw new Error('Protection removal returned null');
      }

      // Check that the protection record was modified
      if (result[10] === 0x01) {
        throw new Error('Protection record was not modified');
      }
    });
  }

  private async testCodeExtraction(): Promise<void> {
    await this.runTest('Enhanced Code Extraction', async () => {
      // Create mock VBA data with module information
      const mockVBAData = new Uint8Array([
        // Convert a simple VBA module structure to bytes
        ...new TextEncoder().encode('Attribute VB_Name = "Module1"\r\n'),
        ...new TextEncoder().encode('Sub TestSub()\r\n'),
        ...new TextEncoder().encode('    MsgBox "Hello World"\r\n'),
        ...new TextEncoder().encode('End Sub\r\n')
      ]);

      const modules = extractVBACodeEnhanced(mockVBAData, this.logger);
      
      if (modules.length === 0) {
        throw new Error('No modules were extracted');
      }

      if (!modules[0].extractionSuccess) {
        throw new Error('Module extraction was marked as failed');
      }

      if (!modules[0].code.includes('TestSub')) {
        throw new Error('Expected VBA code was not found in extracted module');
      }
    });
  }

  private async testErrorHandling(): Promise<void> {
    await this.runTest('Error Handling', async () => {
      // Test with corrupted data
      const corruptedData = new Uint8Array([0xFF, 0xFF, 0xFF, 0xFF]);
      
      const { removeVBAProtectionEnhanced } = await import('../utils/vbaProtectionRemover');
      const result = removeVBAProtectionEnhanced(corruptedData, this.logger);
      
      // Should not crash, should return the original data
      if (!result) {
        throw new Error('Function should handle corrupted data gracefully');
      }

      // Test code extraction with invalid data
      const modules = extractVBACodeEnhanced(corruptedData, this.logger);
      
      // Should return empty array, not crash
      if (!Array.isArray(modules)) {
        throw new Error('Function should return array even for invalid data');
      }
    });
  }

  private async testFileIntegrity(): Promise<void> {
    await this.runTest('File Integrity Validation', async () => {
      // Test that our modifications don't break basic file structure
      
      // Create a minimal valid VBA project structure
      const validProject = new Uint8Array([
        // VBA signature and checksum
        0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        // Mock project data
        ...Array(16).fill(0x00)
      ]);

      // Calculate proper checksum
      let checksum = 0;
      for (let i = 8; i < validProject.length; i++) {
        checksum += validProject[i];
      }
      validProject[4] = checksum & 0xFF;
      validProject[5] = (checksum >> 8) & 0xFF;
      validProject[6] = (checksum >> 16) & 0xFF;
      validProject[7] = (checksum >> 24) & 0xFF;

      // Verify it's valid before modification
      if (!validateVBAProject(validProject, this.logger)) {
        throw new Error('Test project should be valid before modification');
      }

      // Apply protection removal
      const { removeVBAProtectionEnhanced } = await import('../utils/vbaProtectionRemover');
      const modified = removeVBAProtectionEnhanced(validProject, this.logger);

      if (!modified) {
        throw new Error('Protection removal should not return null for valid data');
      }

      // Verify it's still valid after modification
      if (!validateVBAProject(modified, this.logger)) {
        throw new Error('Project should remain valid after protection removal');
      }
    });
  }

  private logSummary(): void {
    const totalTests = this.results.length;
    const passedTests = this.results.filter(r => r.success).length;
    const failedTests = totalTests - passedTests;
    const totalDuration = this.results.reduce((sum, r) => sum + r.duration, 0);

    this.logger(`\n=== Test Summary ===`, 'info');
    this.logger(`Total tests: ${totalTests}`, 'info');
    this.logger(`Passed: ${passedTests}`, passedTests === totalTests ? 'success' : 'info');
    if (failedTests > 0) {
      this.logger(`Failed: ${failedTests}`, 'error');
    }
    this.logger(`Total duration: ${totalDuration}ms`, 'info');
    this.logger(`Success rate: ${((passedTests / totalTests) * 100).toFixed(1)}%`, 
                passedTests === totalTests ? 'success' : 'warning');

    if (failedTests > 0) {
      this.logger('\n=== Failed Tests ===', 'error');
      this.results.filter(r => !r.success).forEach(r => {
        this.logger(`${r.testName}: ${r.message}`, 'error');
      });
    }
  }

  /**
   * Generate a detailed test report
   */
  generateReport(): string {
    let report = 'VBA Toolkit Test Report\n';
    report += '======================\n\n';
    
    report += `Generated: ${new Date().toISOString()}\n`;
    report += `Total Tests: ${this.results.length}\n`;
    report += `Passed: ${this.results.filter(r => r.success).length}\n`;
    report += `Failed: ${this.results.filter(r => !r.success).length}\n\n`;

    this.results.forEach(result => {
      report += `[${result.success ? 'PASS' : 'FAIL'}] ${result.testName}\n`;
      report += `  Duration: ${result.duration}ms\n`;
      if (!result.success) {
        report += `  Error: ${result.message}\n`;
      }
      report += '\n';
    });

    return report;
  }
}