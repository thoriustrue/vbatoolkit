import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { NodeGlobalsPolyfillPlugin } from '@esbuild-plugins/node-globals-polyfill';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    exclude: ['lucide-react'],
    esbuildOptions: {
      // Node.js global to browser global polyfills
      define: {
        global: 'globalThis'
      },
      plugins: [
        NodeGlobalsPolyfillPlugin({
          buffer: true,
          process: true
        })
      ]
    }
  },
  // GitHub Pages configuration for the vbatoolkit repository
  base: '/vbatoolkit/',  // Replace with your actual repository name
  build: {
    // Ensure sourcemaps are generated
    sourcemap: true,
    // Improve output for deployment
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ['react', 'react-dom'],
          xlsx: ['xlsx'],
          jszip: ['jszip']
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
