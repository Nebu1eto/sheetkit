import { type DefaultTheme, defineConfig } from 'vitepress';

export const shared = defineConfig({
  title: 'SheetKit',
  description: 'A Rust library for reading and writing Excel files, with Node.js bindings',

  base: '/sheetkit/',

  lastUpdated: true,
  cleanUrls: true,

  head: [['link', { rel: 'icon', type: 'image/svg+xml', href: '/sheetkit/logo.svg' }]],

  themeConfig: {
    socialLinks: [{ icon: 'github', link: 'https://github.com/Nebu1eto/sheetkit' }],

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
