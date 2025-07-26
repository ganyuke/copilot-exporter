import { defineConfig } from 'vite';
import monkey from 'vite-plugin-monkey';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    monkey({
      entry: 'src/main.ts',
      userscript: {
        icon: 'https://upload.wikimedia.org/wikipedia/commons/0/0e/Microsoft_365_%282022%29.svg',
        namespace: "@ganyuke",
        match: ['https://m365.cloud.microsoft/', 'https://m365.cloud.microsoft/chat/*'],
        'run-at': 'document-end',
      },
    }),
  ],
});
