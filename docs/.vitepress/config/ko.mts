import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { type DefaultTheme, defineConfig } from 'vitepress';
import { fixTypedocSidebarLinks } from './shared.mts';

export const ko = defineConfig({
  lang: 'ko',
  description: 'Rust와 TypeScript를 위한 고성능 SpreadsheetML 라이브러리',

  themeConfig: {
    nav: nav(),
    sidebar: {
      '/ko/guide/': { base: '/ko/guide/', items: sidebarGuide() },
      '/ko/api-reference/': { base: '/ko/api-reference/', items: sidebarApiReference() },
      '/typescript-api/': sidebarTypescriptApi(),
    },
    editLink: {
      pattern: 'https://github.com/Nebu1eto/sheetkit/edit/main/docs/:path',
      text: 'GitHub에서 이 페이지 편집하기',
    },
    footer: {
      message: 'MIT / Apache-2.0 라이선스로 배포됩니다.',
      copyright: 'Copyright 2025 Haze Lee',
    },
  },
});

function nav(): DefaultTheme.NavItem[] {
  return [
    {
      text: '가이드',
      items: [
        { text: '시작 가이드', link: '/ko/getting-started' },
        { text: '아키텍처', link: '/ko/architecture' },
        { text: '성능', link: '/ko/performance' },
        { text: '가이드 개요', link: '/ko/guide/' },
      ],
    },
    { text: 'API 레퍼런스', link: '/ko/api-reference/', activeMatch: '/ko/api-reference/' },
    {
      text: 'API 문서',
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
      text: '소개',
      items: [
        { text: '시작 가이드', link: '/ko/getting-started' },
        { text: '아키텍처', link: '/ko/architecture' },
        { text: '성능', link: '/ko/performance' },
      ],
    },
    {
      text: '가이드',
      items: [
        { text: '개요', link: '/ko/guide/' },
        { text: '기본 작업', link: '/ko/guide/basic-operations' },
        { text: '스타일', link: '/ko/guide/styling' },
        { text: '데이터 기능', link: '/ko/guide/data-features' },
        { text: '렌더링', link: '/ko/guide/rendering' },
        { text: '고급 기능', link: '/ko/guide/advanced' },
        { text: 'CLI', link: '/ko/guide/cli' },
      ],
    },
    {
      text: '커뮤니티',
      items: [{ text: '기여 가이드', link: '/ko/contributing' }],
    },
  ];
}

function sidebarApiReference(): DefaultTheme.SidebarItem[] {
  return [
    {
      text: 'API 레퍼런스',
      items: [
        { text: '개요', link: '/' },
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
