import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  base: '/index/',
  plugins: [vue()],
  server: {
    host: true
  }
})
