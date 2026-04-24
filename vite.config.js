import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueDevTools from 'vite-plugin-vue-devtools'

function inlineBuildAssetsForFileProtocol() {
  const normalizeFileName = (value) => value.replace(/^[./]+/, '')

  return {
    name: 'inline-build-assets-for-file-protocol',
    apply: 'build',
    enforce: 'post',
    generateBundle(_, bundle) {
      const htmlAssets = Object.values(bundle).filter(
        (item) => item.type === 'asset' && item.fileName.endsWith('.html'),
      )

      for (const htmlAsset of htmlAssets) {
        let html = String(htmlAsset.source)
        const removableFiles = new Set()

        html = html.replace(
          /<link[^>]*rel="modulepreload"[^>]*>/g,
          '',
        )

        html = html.replace(
          /<link[^>]*rel="stylesheet"[^>]*href="([^"]+)"[^>]*>/g,
          (tag, href) => {
            const fileName = normalizeFileName(href)
            const asset = bundle[fileName]

            if (!asset || asset.type !== 'asset') {
              return tag
            }

            removableFiles.add(fileName)
            return `<style>\n${String(asset.source)}\n</style>`
          },
        )

        html = html.replace(
          /<script[^>]*type="module"[^>]*src="([^"]+)"[^>]*><\/script>/g,
          (tag, src) => {
            const fileName = normalizeFileName(src)
            const chunk = bundle[fileName]

            if (!chunk || chunk.type !== 'chunk') {
              return tag
            }

            removableFiles.add(fileName)
            return `<script type="module">\n${chunk.code}\n</script>`
          },
        )

        htmlAsset.source = html

        for (const fileName of removableFiles) {
          delete bundle[fileName]
        }
      }
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  base: './',
  build: {
    cssCodeSplit: false,
  },
  plugins: [
    vue(),
    vueDevTools(),
    inlineBuildAssetsForFileProtocol(),
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url))
    },
  },
})
