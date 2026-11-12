import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';
import { readFileSync } from 'node:fs';

const sidebar = JSON.parse(
  readFileSync(new URL('./src/data/sidebar.json', import.meta.url), 'utf8')
);

const site = process.env.DOCS_SITE || 'https://PharmaRA.github.io';
const rawBase = process.env.DOCS_BASE ?? '/RATools-for-Word';
const trimmedBase = rawBase.replace(/^\/+|\/+$/g, '');
const base = trimmedBase ? `/${trimmedBase}` : '/';

export default defineConfig({
  site,
  base,
  trailingSlash: 'always',
  integrations: [
    starlight({
      title: 'RATools for Word',
      description: '面向药品注册资料编写的 Word 效率插件：排版、题注引用、审阅清理与批量文件处理。',
      logo: { src: './src/assets/logo.png', alt: 'RATools for Word' },
      favicon: '/favicon.png',
      defaultLocale: 'root',
      locales: { root: { label: '简体中文', lang: 'zh-CN' } },
      social: [{ icon: 'github', label: 'GitHub', href: 'https://github.com/PharmaRA/RATools-for-Word' }],
      customCss: ['./src/styles/theme.css'],
      components: {
        Header: './src/components/SiteHeader.astro',
      },
      sidebar,
      pagefind: true,
      tableOfContents: { minHeadingLevel: 2, maxHeadingLevel: 3 },
      credits: false,
      head: [
        { tag: 'meta', attrs: { name: 'theme-color', content: '#f7c948' } },
        { tag: 'meta', attrs: { property: 'og:type', content: 'website' } },
        { tag: 'meta', attrs: { property: 'og:locale', content: 'zh_CN' } },
      ],
    }),
  ],
});