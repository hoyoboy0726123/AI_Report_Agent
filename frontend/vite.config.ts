import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// 後端固定跑在 127.0.0.1:8756。
// dev:proxy /api 與 ws 到後端。build:輸出到 backend/static 供 FastAPI 直接服務。
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5275,
    proxy: {
      '/api': {
        target: 'http://127.0.0.1:8756',
        changeOrigin: true,
        ws: true,
      },
    },
  },
  build: {
    outDir: '../backend/static',
    emptyOutDir: true,
  },
})
