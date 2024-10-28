# html-to-pptx

## 描述

在做一个AI PPT的功能时候，找了很久也没找到一个合适的工具可以辅助我完成这件事情， 最后手撸把页面元素转换成ppx，代码很简单，可以源码自行查看。
只实现了一些基础功能，不满足需求可自行修改。 有需要也可在github中提issue。【有时间会更新】

## 安装

```bash
npm install html-to-pptx
```

## 说明


1. 支持转换页面dom元素内容及样式
2. 支持转换页面图片
3. 支持转换页面表格
4. page 页面背景图片

### 使用说明

使用时请遵循以下规定：
- pageClassName 为必填项，为每一页的顶层元素，会解析该元素的背景图片或颜色
- 会自动解析 pageClassName 元素下的所有子元素
- 请在子元素添加 attributes[type] = 'text' | 'image'| 'table', 会对应解析为ppx元素中的文本、图片、表格
- dom元素请使用flex布局， 对应应用至文本的对齐方式，规则如下:
```typescript
/** 垂直对齐方式转换为 pptx 的 valign */
function alignItemsToPPTvalign(val: string): string {
  switch (val) {
    case 'flex-start':
    case 'top':
      return 'top'; // PPTX 中对应于顶部对齐
    case 'center':
      return 'middle'; // PPTX 中对应于垂直居中
    case 'flex-end':
    case 'bottom':
      return 'bottom'; // PPTX 中对应于底部对齐
    case 'stretch':
      return 'middle'; // 默认为中间，可以根据需求调整
    default:
      return 'middle'; // 默认返回中间对齐
  }
}
/** 水平对齐方式转换为 pptx 的 align */
function justifyContentToPPTalign(val: string): string {
  switch (val) {
    case 'flex-start':
    case 'left':
      return 'left'; // PPTX 中对应于左对齐
    case 'center':
      return 'center'; // PPTX 中对应于水平居中
    case 'flex-end':
    case 'right':
      return 'right'; // PPTX 中对应于右对齐
    case 'space-between':
      return 'justify'; // PPTX 中的两端对齐
    case 'space-around':
    case 'space-evenly':
      return 'justify'; // PPTX 中的均匀对齐
    default:
      return 'left'; // 默认返回左对齐
  }
}
```

## 注意

- 暂不支持以下内容：
1. chat 表格
- 转换的样式并不全面， 暂时已转换的样式如下列表
| CSS 属性        | 描述              |
|----------------|-------------------|
| `width`        | 元素的宽度        |
| `height`       | 元素的高度        |
| `top`          | 元素的上边距      |
| `left`         | 元素的左边距      |
| `color`        | 文本颜色          |
| `align-items`  | Flexbox 对齐方式  |
| `justify-content` | Flexbox 主轴对齐方式 |
| `font-size`    | 字体大小          |
| `font-weight`  | 字体粗细          |
| `border-color` | 边框颜色          |


