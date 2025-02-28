import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { NodeGlobalsPolyfillPlugin } from '@esbuild-plugins/node-globals-polyfill';
import { NodeModulesPolyfillPlugin } from '@esbuild-plugins/node-modules-polyfill';

/**
 * Vite configuration for the Excel VBA Toolkit
 * 
 * This configuration includes:
 * - React plugin for JSX/TSX support
 * - Node.js polyfills for browser compatibility
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
    include: ['buffer', 'xlsx/dist/xlsx.full.min.js'],
    esbuildOptions: {
      define: {
        global: 'globalThis'
      },
      plugins: [
        NodeGlobalsPolyfillPlugin({
          buffer: true,
          process: true
        }),
        NodeModulesPolyfillPlugin()
      ]
    }
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
