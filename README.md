# HTML to PPTX

一个强大的前端库，用于将 HTML DOM 元素转换为可编辑的 PowerPoint (.pptx) 文件。

它不仅仅是截图，而是将 HTML 元素解析为**原生的 PPT 形状、文本框、表格和图表**，因此生成的 PPT 是**清晰且可编辑**的。

## ✨ 特性

- **所见即所得**：基于 DOM 的 `getBoundingClientRect` 计算坐标，高度还原网页布局。
- **可编辑文本**：HTML 文本会被转换为独立的 PPT 文本框，支持字体、颜色、字号、加粗、斜体、下划线等样式。
- **原生表格**：`<table>` 标签会被转换为 PPT 的原生表格，支持合并单元格 (`rowspan`/`colspan`)。
- **原生图表支持**：通过 `data-` 属性配置，可以将 HTML 元素渲染为可编辑的 PPT 图表。
- **动画支持**：支持 CSS 动画（如 `fadeIn`, `slideInUp`）映射到 PPT 动画，也支持自定义动画属性。
- **样式还原**：支持背景色、背景图、边框、圆角（自动识别矩形/圆角矩形/椭圆）、透明度、旋转角度等。
- **多页支持**：通过指定类名，自动分割多页 PPT。
- **Canvas & 图片**：自动处理 `<img>` 和 `<canvas>` 元素。

## 📦 安装

```bash
npm install html-to-pptx
# 或者
yarn add html-to-pptx
```

> 注意：本项目依赖 `pptxgenjs` 作为底层生成引擎。

## 🚀 快速开始

### 1. 准备 HTML 结构

在你的页面中，将每一页 PPT 的内容包裹在一个容器中，并赋予一个统一的类名（例如 `h-ppt-page`）。

```html
<!-- 第一页 -->
<div
  class="h-ppt-page"
  style="width: 1000px; height: 562.5px; position: relative; background: white;"
>
  <h1 style="color: red;">Hello World</h1>
  <p style="font-size: 20px;">这是一个示例文本。</p>
</div>

<!-- 第二页 -->
<div
  class="h-ppt-page"
  style="width: 1000px; height: 562.5px; position: relative;"
>
  <img src="image.png" />
</div>
```

> **建议**：为了获得最佳效果，建议将容器宽高比设置为 **16:9** (例如 1000px \* 562.5px)，因为插件默认输出 16:9 的幻灯片。

### 2. 调用导出函数

在你的 JavaScript/TypeScript 代码中：

```typescript
import { downloadHtmlToPpt } from "html-to-pptx";

const handleExport = async () => {
  try {
    // 传入类名和文件名
    await downloadHtmlToPpt("h-ppt-page", "我的演示文稿");
    console.log("下载成功！");
  } catch (error) {
    console.error("导出失败", error);
  }
};
```

## 📚 API 文档

### `downloadHtmlToPpt(pageClassName, fileName)`

直接在浏览器中触发 PPT 文件下载。

| 参数            | 类型            | 默认值           | 描述                                |
| :-------------- | :-------------- | :--------------- | :---------------------------------- |
| `pageClassName` | `string`        | `"page"`         | 标识每一页 PPT 的 DOM 类名          |
| `fileName`      | `string`        | `"presentation"` | 下载的文件名（不需要加 .pptx 后缀） |
| **返回**        | `Promise<void>` | -                | -                                   |

### `exportHtmlToPpt(pageClassName, outputType)`

生成 PPT 数据流，不自动下载，适合需要上传到服务器或自定义处理的场景。

| 参数            | 类型                             | 默认值   | 描述                                                               |
| :-------------- | :------------------------------- | :------- | :----------------------------------------------------------------- |
| `pageClassName` | `string`                         | `"page"` | 标识每一页 PPT 的 DOM 类名                                         |
| `outputType`    | `string`                         | `"blob"` | 输出类型: `"arraybuffer"`, `"base64"`, `"blob"`, `"uint8array"` 等 |
| **返回**        | `Promise<Blob \| string \| ...>` | -        | 返回对应格式的文件数据                                             |

## 🎨 高级功能

### 1. 原生图表 (Native Charts)

你可以通过在 DOM 元素上添加 `data-pptx-chart-config` 属性，让插件在 PPT 中生成可编辑的图表，而不是截图。

该属性的值必须是一个符合 JSON 格式的字符串，包含 `type`, `data`, 和 `options`。

```html
<div
  style="width: 400px; height: 300px;"
  data-pptx-chart-config='{
    "type": "bar",
    "data": [
      {
        "name": "Series 1",
        "labels": ["Category 1", "Category 2"],
        "values": [10, 20]
      }
    ],
    "options": {
      "chartColors": ["FF0000"]
    }
  }'
>
  <!-- 这里可以放一个普通的 img 或 canvas 用于网页展示，PPT 导出时会被替换为原生图表 -->
</div>
```

### 2. 动画支持 (Animations)

插件支持两种方式添加动画：

**方式 A: 自动识别 CSS 动画**
插件会自动解析元素的 `animation-name` 并映射到 PPT 动画。
支持的 CSS 动画名包括：`fadeIn`, `fadeInUp`, `slideInUp`, `zoomIn`, `bounceIn` 等。

**方式 B: 使用 `data-` 属性（推荐）**
显式指定 PPT 动画类型、持续时间和延迟。

```html
<div
  data-pptx-animation="flyInBottom"
  data-pptx-duration="1500"
  data-pptx-delay="500"
>
  从底部飞入的文本
</div>
```

- `data-pptx-animation`: 动画类型 (参考 PptxGenJS 文档)
- `data-pptx-duration`: 持续时间 (毫秒)
- `data-pptx-delay`: 延迟时间 (毫秒)

### 3. 忽略元素

如果你希望某个元素在网页显示，但在导出 PPT 时被忽略，可以添加 `data-html2pptx-ignore` 属性（需在样式中配合 `display: none` 或在代码逻辑中自行扩展，目前核心代码逻辑会忽略 `display: none`, `visibility: hidden`, `opacity: 0` 的元素）。

## ⚠️ 注意事项 & 限制

1.  **CORS 跨域图片**：如果页面包含跨域图片（`<img>`），请确保服务器配置了正确的 CORS 头，或者图片允许跨域访问，否则可能无法生成到 PPT 中。
2.  **布局依赖**：插件通过 `getBoundingClientRect` 计算位置。请确保在调用导出函数时，DOM 元素是**可见的**且**已渲染**。如果在 `display: none` 的容器中进行转换，坐标计算将失败。
3.  **CSS 支持**：虽然支持大部分常用 CSS，但复杂的 CSS 布局（如复杂的伪元素、复杂的 `clip-path`、滤镜等）可能无法完美还原，建议使用简单的 Flex/Grid/Absolute 布局。
4.  **字体**：PPT 文件中引用的字体取决于用户系统是否安装了该字体。插件默认设置了 fallback 字体。

## 🛠 开发与贡献

欢迎提交 Issue 或 Pull Request 来改进此项目！

## 📄 License

MIT
