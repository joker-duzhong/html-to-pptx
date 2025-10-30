import typescript from "rollup-plugin-typescript2";
import { terser } from "rollup-plugin-terser";
import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default {
  input: "src/index.ts",
  output: [
    {
      file: "dist/index.umd.js",
      format: "umd",
      name: "HtmlToPptx",
      sourcemap: false,
      exports: 'named'
    },
    {
      file: "dist/index.esm.js",
      format: "es",
      sourcemap: false
    }
  ],
  plugins: [
    resolve(),
    commonjs(),
    typescript(),
    terser()
  ],
};
