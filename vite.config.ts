// defineConfig comes from 'vitest/config' (a superset of vite's) so the
// `test` block below is type-checked correctly.
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  test: {
    // jsdom gives us DOMParser/XMLSerializer for OOXML unit tests without a
    // real browser. happy-dom is lighter but its getElementsByTagNameNS is
    // broken for XML (non-HTML) documents, which every OOXML helper here
    // depends on - jsdom's XML DOM support is complete enough to rely on.
    environment: 'jsdom',
  },
})
