import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

/**
 * Vite configuration for the Excel VBA Toolkit
 * 
 * This configuration includes:
 * - React plugin for JSX/TSX support
 * - Buffer polyfill for binary data handling
 * - GitHub Pages deployment settings
 */
export default defineConfig({
  plugins: [react()],
  
  // Define global values
  define: {
    global: 'globalThis',
  },
  
  // Optimize dependencies
  optimizeDeps: {
    include: ['buffer'],
    exclude: ['@esbuild-plugins/node-globals-polyfill', '@esbuild-plugins/node-modules-polyfill']
  },
  
  // GitHub Pages configuration
  base: '/vbatoolkit/',
  
  // Build configuration
  build: {
    sourcemap: true,
    assetsDir: 'assets',
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ['react', 'react-dom'],
          xlsx: ['xlsx'],
          jszip: ['jszip'],
          buffer: ['buffer']
        }
      }
    }
  },
  
  // Development server
  server: {
    port: 3000,
    open: true
  }
});
