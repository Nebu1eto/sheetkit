import { type DefaultTheme, defineConfig } from 'vitepress';

export const shared = defineConfig({
  title: 'SheetKit',
  description: 'High-performance SpreadsheetML library for Rust and TypeScript',

  base: '/sheetkit/',

  lastUpdated: true,
  cleanUrls: true,

  head: [
    ['link', { rel: 'icon', type: 'image/svg+xml', href: '/sheetkit/logo.svg' }],
    ['link', { rel: 'icon', type: 'image/png', sizes: '32x32', href: '/sheetkit/favicon-32x32.png' }],
    ['link', { rel: 'icon', type: 'image/png', sizes: '16x16', href: '/sheetkit/favicon-16x16.png' }],
    ['link', { rel: 'apple-touch-icon', sizes: '180x180', href: '/sheetkit/apple-touch-icon.png' }],
  ],

  themeConfig: {
    logo: '/logo.svg',
    socialLinks: [
      { icon: 'github', link: 'https://github.com/Nebu1eto/sheetkit' },
      { icon: 'npm', link: 'https://www.npmjs.com/package/@sheetkit/node' },
      { icon: 'rust', link: 'https://crates.io/crates/sheetkit' },
    ],

    search: {
      provider: 'local',
    },
  },
});

export function fixTypedocSidebarLinks(
  items: DefaultTheme.SidebarItem[],
): DefaultTheme.SidebarItem[] {
  return items.map((item) => ({
    ...item,
    link: item.link?.replace(/^\/docs\//, '/').replace(/\.md$/, ''),
    items: item.items ? fixTypedocSidebarLinks(item.items) : undefined,
  }));
}
