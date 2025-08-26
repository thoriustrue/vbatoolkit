/**
 * Demo component to test and showcase the enhanced VBA toolkit functionality
 */

import React, { useState } from 'react';
import { VBAToolkitTester } from '../utils/vbaToolkitTester';
import { validateVBAProject } from '../utils/vbaProtectionRemover';
import { extractVBACodeEnhanced } from '../utils/vbaCodeExtractor/enhancedExtractor';

interface DemoLog {
  message: string;
  type: 'info' | 'success' | 'warning' | 'error';
  timestamp: Date;
}

export function VBAToolkitDemo() {
  const [logs, setLogs] = useState<DemoLog[]>([]);
  const [isRunning, setIsRunning] = useState(false);

  const addLog = (message: string, type: 'info' | 'success' | 'warning' | 'error') => {
    setLogs(prev => [...prev, { message, type, timestamp: new Date() }]);
  };

  const clearLogs = () => {
    setLogs([]);
  };

  const runBasicTests = async () => {
    setIsRunning(true);
    clearLogs();
    
    addLog('Starting VBA Toolkit Enhanced Functionality Tests', 'info');
    
    try {
      // Test 1: VBA Validation
      addLog('=== Test 1: VBA Project Validation ===', 'info');
      
      // Create a mock VBA project with proper checksum
      const mockVBA = new Uint8Array(24);
      mockVBA[0] = 0x01; // Mock signature
      mockVBA[1] = 0x00;
      
      // Add some data
      for (let i = 8; i < mockVBA.length; i++) {
        mockVBA[i] = i - 7; // Simple test pattern
      }
      
      // Calculate checksum
      let checksum = 0;
      for (let i = 8; i < mockVBA.length; i++) {
        checksum += mockVBA[i];
      }
      mockVBA[4] = checksum & 0xFF;
      mockVBA[5] = (checksum >> 8) & 0xFF;
      mockVBA[6] = (checksum >> 16) & 0xFF;
      mockVBA[7] = (checksum >> 24) & 0xFF;
      
      const isValid = validateVBAProject(mockVBA, addLog);
      addLog(`VBA validation result: ${isValid ? 'VALID' : 'INVALID'}`, isValid ? 'success' : 'error');
      
      // Test 2: Enhanced Code Extraction  
      addLog('=== Test 2: Enhanced Code Extraction ===', 'info');
      
      // Create mock VBA code
      const mockCode = `Attribute VB_Name = "TestModule"
Attribute VB_Type = 0

Sub HelloWorld()
    MsgBox "Hello from enhanced VBA extraction!"
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function`;
      
      const mockVBACode = new TextEncoder().encode(mockCode);
      const extractedModules = extractVBACodeEnhanced(mockVBACode, addLog);
      
      addLog(`Extracted ${extractedModules.length} modules`, extractedModules.length > 0 ? 'success' : 'warning');
      
      extractedModules.forEach((module, index) => {
        addLog(`Module ${index + 1}: ${module.name} (${module.type}) - ${module.extractionSuccess ? 'SUCCESS' : 'FAILED'}`, 
               module.extractionSuccess ? 'success' : 'warning');
        if (module.extractionSuccess) {
          addLog(`  Code lines: ${module.code.split('\n').length}`, 'info');
        }
      });
      
      // Test 3: Comprehensive Test Suite
      addLog('=== Test 3: Running Comprehensive Test Suite ===', 'info');
      
      const tester = new VBAToolkitTester(addLog);
      const results = await tester.runTests();
      
      const passed = results.filter(r => r.success).length;
      const total = results.length;
      
      addLog(`Test suite completed: ${passed}/${total} tests passed`, 
             passed === total ? 'success' : 'warning');
      
      // Test 4: Performance Test
      addLog('=== Test 4: Performance Validation ===', 'info');
      
      const startTime = Date.now();
      
      // Simulate processing a larger VBA project
      const largeVBA = new Uint8Array(10000);
      for (let i = 0; i < largeVBA.length; i++) {
        largeVBA[i] = Math.floor(Math.random() * 256);
      }
      
      const largeModules = extractVBACodeEnhanced(largeVBA, addLog);
      const endTime = Date.now();
      
      addLog(`Processed ${largeVBA.length} bytes in ${endTime - startTime}ms`, 'info');
      addLog(`Found ${largeModules.length} potential modules`, 'info');
      
      addLog('=== All Tests Completed ===', 'success');
      
    } catch (error) {
      addLog(`Error during testing: ${error instanceof Error ? error.message : String(error)}`, 'error');
    } finally {
      setIsRunning(false);
    }
  };

  const generateReport = async () => {
    addLog('Generating comprehensive test report...', 'info');
    
    const tester = new VBAToolkitTester(addLog);
    await tester.runTests();
    const report = tester.generateReport();
    
    // Create and download the report
    const blob = new Blob([report], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `vba-toolkit-test-report-${new Date().toISOString().split('T')[0]}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    addLog('Test report downloaded successfully', 'success');
  };

  const getLogColor = (type: string) => {
    switch (type) {
      case 'success': return 'text-green-600';
      case 'error': return 'text-red-600';
      case 'warning': return 'text-yellow-600';
      default: return 'text-gray-700';
    }
  };

  return (
    <div className="max-w-4xl mx-auto p-6">
      <div className="bg-white rounded-lg shadow-lg p-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-4">
          VBA Toolkit Enhanced Functionality Demo
        </h2>
        
        <p className="text-gray-600 mb-6">
          This demo showcases the enhanced VBA processing capabilities including proper binary format parsing,
          improved protection removal, and robust code extraction.
        </p>
        
        <div className="flex gap-4 mb-6">
          <button
            onClick={runBasicTests}
            disabled={isRunning}
            className={`px-4 py-2 rounded-md font-medium ${
              isRunning 
                ? 'bg-gray-400 cursor-not-allowed text-white' 
                : 'bg-blue-600 hover:bg-blue-700 text-white'
            }`}
          >
            {isRunning ? 'Running Tests...' : 'Run Enhancement Tests'}
          </button>
          
          <button
            onClick={generateReport}
            disabled={isRunning}
            className="px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-md font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            Generate Test Report
          </button>
          
          <button
            onClick={clearLogs}
            className="px-4 py-2 bg-gray-600 hover:bg-gray-700 text-white rounded-md font-medium"
          >
            Clear Logs
          </button>
        </div>
        
        <div className="bg-gray-50 rounded-lg p-4 max-h-96 overflow-y-auto">
          <h3 className="text-lg font-semibold mb-2">Test Results & Logs</h3>
          
          {logs.length === 0 ? (
            <p className="text-gray-500 italic">No logs yet. Run the tests to see the enhanced functionality in action.</p>
          ) : (
            <div className="space-y-1 font-mono text-sm">
              {logs.map((log, index) => (
                <div key={index} className={`${getLogColor(log.type)}`}>
                  <span className="text-gray-500 text-xs">
                    [{log.timestamp.toLocaleTimeString()}]
                  </span>{' '}
                  {log.message}
                </div>
              ))}
            </div>
          )}
        </div>
        
        <div className="mt-6 p-4 bg-blue-50 rounded-lg">
          <h3 className="text-lg font-semibold text-blue-900 mb-2">Enhancement Highlights</h3>
          <ul className="text-blue-800 space-y-1">
            <li>• <strong>Enhanced VBA Protection Removal:</strong> Proper VBA record parsing instead of basic pattern matching</li>
            <li>• <strong>Improved Code Extraction:</strong> Direct binary parsing with decompression support</li>
            <li>• <strong>Validation Framework:</strong> Checksum and structure validation at each step</li>
            <li>• <strong>Comprehensive Testing:</strong> Automated test suite to validate functionality</li>
            <li>• <strong>Graceful Fallbacks:</strong> Enhanced methods with legacy support for compatibility</li>
          </ul>
        </div>
      </div>
    </div>
  );
}