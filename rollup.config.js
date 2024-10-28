import typescript from "rollup-plugin-typescript2";
import { terser } from "rollup-plugin-terser";

export default {
  input: "src/index.ts", // 入口文件
  output: {
    file: "dist/index.js", // 输出文件
    format: "cjs", // 输出格式，可以根据需要选择 'cjs', 'esm', 'umd' 等
    sourcemap: false, // 是否生成 source map
  },
  plugins: [
    typescript(), // 编译 TypeScript
    terser(), // 压缩代码
  ],
};
