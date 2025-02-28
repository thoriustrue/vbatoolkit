import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { NodeGlobalsPolyfillPlugin } from '@esbuild-plugins/node-globals-polyfill';
import { NodeModulesPolyfillPlugin } from '@esbuild-plugins/node-modules-polyfill';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Define global values
  define: {
    global: 'globalThis',
    // Ensure Buffer is defined for library code
    'global.Buffer': 'globalThis.Buffer',
  },
  optimizeDeps: {
    include: [
      'buffer',
      'xlsx/dist/xlsx.full.min.js' // Force include XLSX
    ],
    esbuildOptions: {
      // Node.js global to browser global polyfills
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
  // GitHub Pages configuration for the vbatoolkit repository
  base: '/vbatoolkit/',  // Replace with your actual repository name
  build: {
    // Ensure sourcemaps are generated
    sourcemap: true,
    // Improve output for deployment
    assetsInlineLimit: 0,
    rollupOptions: {
      output: {
        entryFileNames: `assets/[name].js`,
        chunkFileNames: `assets/[name].js`,
        assetFileNames: `assets/[name].[ext]`,
        manualChunks: {
          vendor: ['react', 'react-dom'],
          xlsx: ['xlsx'],
          jszip: ['jszip'],
          // Create a separate chunk for buffer polyfill
          polyfill: ['buffer']
        }
      }
    }
  },
  server: {
    cors: {
      origin: ['http://localhost:3000', 'https://thoriustrue.github.io'],
      methods: ['GET', 'POST']
    }
  }
});
