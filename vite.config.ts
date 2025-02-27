import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    exclude: ['lucide-react'],
  },
  // GitHub Pages configuration for the vbatoolkit repository
  base: '/vbatoolkit/',  // Explicitly use the repository name
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
  }
});
