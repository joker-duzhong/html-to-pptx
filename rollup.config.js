import resolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import typescript from "@rollup/plugin-typescript";
import { terser } from "rollup-plugin-terser";
import pkg from "./package.json";

export default {
  input: "src/index.ts", // 你的入口文件
  output: [
    // 1. CommonJS (给 Node.js 或旧版构建工具用)
    {
      file: pkg.main,
      format: "cjs",
      sourcemap: true,
    },
    // 2. ES Module (给 Vite, Webpack, Rollup 用)
    {
      file: pkg.module,
      format: "esm",
      sourcemap: true,
    },
    // 3. UMD (给浏览器 <script> 标签用)
    {
      file: "dist/html-to-pptx.min.js",
      format: "umd",
      name: "HtmlToPptx", // 全局变量名，浏览器引入后通过 window.HtmlToPptx 访问
      sourcemap: true,
      globals: {
        pptxgenjs: "PptxGenJS", // 告诉 Rollup，pptxgenjs 在全局变量里叫 PptxGenJS
      },
      plugins: [terser()], // 压缩代码
    },
  ],
  // 将 pptxgenjs 视为外部依赖，不打包进你的库中，减小体积
  external: ["pptxgenjs"],
  plugins: [
    resolve(),
    commonjs(),
    typescript({
      tsconfig: "./tsconfig.json",
    }),
  ],
};
