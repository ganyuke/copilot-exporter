import { defineConfig } from 'vite';
import monkey from 'vite-plugin-monkey';
import packageJson from './package.json'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    monkey({
      entry: 'src/main.ts',
      userscript: {
        icon: 'https://upload.wikimedia.org/wikipedia/commons/0/0e/Microsoft_365_%282022%29.svg',
        namespace: packageJson.author.name,
        name: packageJson.title,
        match: ['https://m365.cloud.microsoft/', 'https://m365.cloud.microsoft/chat/*'],
        'run-at': 'document-end'
      },
      build: {
        fileName: `${packageJson.name.split("/").pop()}.user.js`
      }
    }),
  ],
});
