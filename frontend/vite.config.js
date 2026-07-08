import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',      // <-- Add this
    port: 4100,
    strictPort: true,
    proxy: {
      '/generate-sow': 'http://localhost:8000',
      '/health':       'http://localhost:8000',
    }
  }
})
