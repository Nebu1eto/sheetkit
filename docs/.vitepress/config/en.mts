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
      '/guide/': sidebarGuide(),
      '/api-reference/': sidebarApiReference(),
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
        { text: 'Async-First Migration (v0.5+)', link: '/guide/migration-async-first' },
        { text: 'CLI', link: '/guide/cli' },
      ],
    },
    {
      text: 'Community',
      items: [
        { text: 'Contributing', link: '/contributing' },
        {
          text: 'Crates.io',
          link: 'https://crates.io/crates/sheetkit',
          target: '_blank',
        },
        {
          text: 'npm',
          link: 'https://www.npmjs.com/package/@sheetkit/node',
          target: '_blank',
        },
      ],
    },
  ];
}

function sidebarApiReference(): DefaultTheme.SidebarItem[] {
  return [
    {
      text: 'API Reference',
      items: [
        { text: 'Overview', link: '/api-reference/' },
        { text: 'Workbook', link: '/api-reference/workbook' },
        { text: 'Sheet', link: '/api-reference/sheet' },
        { text: 'Cell', link: '/api-reference/cell' },
        { text: 'Row & Column', link: '/api-reference/row-column' },
        { text: 'Style', link: '/api-reference/style' },
        { text: 'Chart', link: '/api-reference/chart' },
        { text: 'Image', link: '/api-reference/image' },
        { text: 'Shape', link: '/api-reference/shape' },
        { text: 'Slicer', link: '/api-reference/slicer' },
        { text: 'Form Control', link: '/api-reference/form-control' },
        { text: 'Data Features', link: '/api-reference/data-features' },
        { text: 'Advanced', link: '/api-reference/advanced' },
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
