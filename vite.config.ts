import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    exclude: ['lucide-react'],
  },
  // Add base configuration for GitHub Pages - CRITICAL!
  // This should match your GitHub repository name
  base: './',  // Use relative paths for better compatibility
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
