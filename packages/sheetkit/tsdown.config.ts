import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['index.ts', 'buffer-codec.ts', 'sheet-data.ts'],
  outDir: '.',
  format: 'esm',
  dts: true,
  unbundle: true,
  clean: false,
  fixedExtension: false,
  sourcemap: false,
  external: [/\.\/binding/],
});
