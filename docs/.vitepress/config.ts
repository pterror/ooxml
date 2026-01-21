import { defineConfig } from 'vitepress'
import { withMermaid } from 'vitepress-plugin-mermaid'

export default withMermaid(
  defineConfig({
    vite: {
      optimizeDeps: {
        include: ['mermaid'],
      },
    },
    title: 'ooxml',
    description: 'Rust library for Office Open XML formats',

    base: '/ooxml/',

    themeConfig: {
      nav: [
        { text: 'Guide', link: '/guide/' },
        { text: 'API', link: '/api/' },
        { text: 'ADRs', link: '/adr/' },
        { text: 'Rhizome', link: 'https://rhizome-lab.github.io/' },
      ],
      sidebar: {
        '/guide/': [
          {
            text: 'Getting Started',
            items: [
              { text: 'Introduction', link: '/guide/' },
              { text: 'Installation', link: '/guide/installation' },
            ],
          },
        ],
        '/adr/': [
          {
            text: 'Architecture Decisions',
            items: [
              { text: 'Overview', link: '/adr/' },
              { text: '001: Custom RNC Parser', link: '/adr/001-custom-rnc-parser' },
            ],
          },
        ],
      },
      socialLinks: [
        { icon: 'github', link: 'https://github.com/pterror/ooxml' },
      ],
      search: {
        provider: 'local',
      },
      editLink: {
        pattern: 'https://github.com/pterror/ooxml/edit/master/docs/:path',
        text: 'Edit this page on GitHub',
      },
    },
  })
)
