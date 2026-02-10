import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { type DefaultTheme, defineConfig } from 'vitepress';
import { fixTypedocSidebarLinks } from './shared.mts';

export const en = defineConfig({
  lang: 'en',
  description: 'High-performance SpreadsheetML library for Rust and TypeScript',

  themeConfig: {
    nav: nav(),
    sidebar: {
      '/guide/': { base: '/guide/', items: sidebarGuide() },
      '/api-reference/': { base: '/api-reference/', items: sidebarApiReference() },
      '/typescript-api/': sidebarTypescriptApi(),
    },
    editLink: {
      pattern: 'https://github.com/Nebu1eto/sheetkit/edit/main/docs/:path',
      text: 'Edit this page on GitHub',
    },
    footer: {
      message: 'Released under the MIT / Apache-2.0 License.',
      copyright: 'Copyright 2025 Haze Lee',
    },
  },
});

function nav(): DefaultTheme.NavItem[] {
  return [
    {
      text: 'Guide',
      items: [
        { text: 'Getting Started', link: '/getting-started' },
        { text: 'Architecture', link: '/architecture' },
        { text: 'Performance', link: '/performance' },
        { text: 'Guide Overview', link: '/guide/' },
      ],
    },
    { text: 'API Reference', link: '/api-reference/', activeMatch: '/api-reference/' },
    {
      text: 'API Docs',
      items: [
        { text: 'Rust (rustdoc)', link: '/rustdoc/sheetkit/index.html', target: '_blank' },
        { text: 'TypeScript (TypeDoc)', link: '/typescript-api/' },
      ],
    },
  ];
}

function sidebarGuide(): DefaultTheme.SidebarItem[] {
  return [
    {
      text: 'Introduction',
      items: [
        { text: 'Getting Started', link: '/getting-started' },
        { text: 'Architecture', link: '/architecture' },
        { text: 'Performance', link: '/performance' },
      ],
    },
    {
      text: 'Guide',
      items: [
        { text: 'Overview', link: '/guide/' },
        { text: 'Basic Operations', link: '/guide/basic-operations' },
        { text: 'Styling', link: '/guide/styling' },
        { text: 'Data Features', link: '/guide/data-features' },
        { text: 'Rendering', link: '/guide/rendering' },
        { text: 'Advanced', link: '/guide/advanced' },
        { text: 'CLI', link: '/guide/cli' },
      ],
    },
    {
      text: 'Community',
      items: [{ text: 'Contributing', link: '/contributing' }],
    },
  ];
}

function sidebarApiReference(): DefaultTheme.SidebarItem[] {
  return [
    {
      text: 'API Reference',
      items: [
        { text: 'Overview', link: '/' },
        { text: 'Workbook', link: '/workbook' },
        { text: 'Sheet', link: '/sheet' },
        { text: 'Cell', link: '/cell' },
        { text: 'Row & Column', link: '/row-column' },
        { text: 'Style', link: '/style' },
        { text: 'Chart', link: '/chart' },
        { text: 'Image', link: '/image' },
        { text: 'Shape', link: '/shape' },
        { text: 'Slicer', link: '/slicer' },
        { text: 'Form Control', link: '/form-control' },
        { text: 'Data Features', link: '/data-features' },
        { text: 'Advanced', link: '/advanced' },
      ],
    },
  ];
}

function sidebarTypescriptApi(): DefaultTheme.SidebarItem[] {
  try {
    const sidebarPath = resolve(import.meta.dirname, '../../typescript-api/typedoc-sidebar.json');
    const raw = JSON.parse(readFileSync(sidebarPath, 'utf-8'));
    return fixTypedocSidebarLinks(raw);
  } catch {
    return [{ text: 'TypeScript API', items: [{ text: 'Overview', link: '/typescript-api/' }] }];
  }
}
