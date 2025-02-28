'use strict';

// Use Uint8Array instead of Buffer
const ZIP_SIGNATURE = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

function isValidZip(outputFileBuffer) {
  // Handle different input types
  let headerBuffer;
  if (outputFileBuffer instanceof ArrayBuffer) {
    headerBuffer = new Uint8Array(outputFileBuffer.slice(0, 4));
  } else if (outputFileBuffer instanceof Uint8Array) {
    headerBuffer = outputFileBuffer.slice(0, 4);
  } else {
    // Fallback for other types
    return false;
  }

  // Compare arrays
  if (headerBuffer.length !== ZIP_SIGNATURE.length) return false;
  for (let i = 0; i < headerBuffer.length; i++) {
    if (headerBuffer[i] !== ZIP_SIGNATURE[i]) return false;
  }
  return true;
}

// Export using ES module syntax for Vite/React compatibility
export { isValidZip };
export default { isValidZip };

// Also support CommonJS for any build tools that might use it
if (typeof exports !== 'undefined') {
  exports.isValidZip = isValidZip;
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { isValidZip };
  }
} 