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
    mermaid: {},
  }),
);
