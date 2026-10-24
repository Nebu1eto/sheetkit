import { defineConfig } from 'vitepress';
import llmstxt from 'vitepress-plugin-llms';
import { withMermaid } from 'vitepress-plugin-mermaid';
import { en } from './config/en.mts';
import { ko } from './config/ko.mts';
import { shared } from './config/shared.mts';

export default withMermaid(
  defineConfig({
    ...shared,
    locales: {
      root: { label: 'English', ...en },
      ko: { label: '한국어', ...ko },
    },
    mermaid: {
      startOnLoad: false,
      flowchart: {
        useMaxWidth: true,
        htmlLabels: true,
        wrappingWidth: 140,
        nodeSpacing: 28,
        rankSpacing: 44,
        curve: 'linear',
      },
      themeVariables: {
        fontSize: '16px',
        edgeLabelBackground: 'transparent',
      },
    },
    vite: {
      build: {
        // The local search index (about 1.1 MB for the English docs) and the mermaid
        // layout engines are single chunks by construction; keep the limit above them.
        chunkSizeWarningLimit: 1200,
      },
      plugins: [
        llmstxt({
          ignoreFiles: ['README.md'],
        }),
      ],
    },
  }),
);
