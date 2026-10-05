import { defineConfig } from 'vite';
import path from 'node:path';
export default defineConfig({
  resolve: { dedupe: ['react', 'react-dom', '@fluentui/react-components', '@griffel/react'] },
  server: { fs: { allow: [path.resolve(__dirname, '..')] }, proxy: { '/api': {target:'http://127.0.0.1:5174', changeOrigin:false}, '/__mock': 'http://127.0.0.1:5174', '/mock-sharepoint': 'http://127.0.0.1:5174' } },
  build: { outDir: 'dist', rollupOptions: { onwarn(warning, warn) { if (warning.code !== 'MODULE_LEVEL_DIRECTIVE') warn(warning); } } },
});
