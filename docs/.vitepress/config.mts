import { defineConfig } from 'vitepress';
import { en } from './config/en.mts';
import { ko } from './config/ko.mts';
import { shared } from './config/shared.mts';

export default defineConfig({
  ...shared,
  locales: {
    root: { label: 'English', ...en },
    ko: { label: '한국어', ...ko },
  },
});
