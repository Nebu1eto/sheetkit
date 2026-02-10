import { defineConfig } from 'vitepress';
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
  }),
);
