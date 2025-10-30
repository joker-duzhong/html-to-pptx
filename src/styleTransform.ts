/**
 * HTML/CSS 到 PPTX 样式转换工具
 * 
 * 不支持的 CSS 属性（PPTX 限制）：
 * - transform: PPTX 不支持 CSS 变换（rotate、scale、translate 等）
 * - opacity: PPTX 不支持整体透明度（但支持颜色中的 alpha 通道）
 * - box-shadow: PPTX 不支持阴影效果
 * - border-radius: PPTX 不支持圆角
 * - background-clip: PPTX 不支持背景裁剪
 * - filter: PPTX 不支持滤镜效果
 * - animation: PPTX 不支持 CSS 动画
 * - flex/grid: 需要在转换时计算出具体的位置和尺寸
 * 
 * 部分支持的属性：
 * - font-family: 仅支持已安装的字体
 * - line-height: 转换为近似值
 * - text-decoration: 仅支持 underline
 */

import PptxGenJS from 'pptxgenjs';

/** 基础样式属性集合 */
export interface TransformedStyle {
  // 尺寸与位置
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  
  // 边距
  margin?: [number, number, number, number]; // [top, right, bottom, left]
  padding?: [number, number, number, number];
  
  // 文本样式
  fontSize?: number;
  fontFace?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: any;
  
  // 对齐
  align?: 'left' | 'center' | 'right' | 'justify';
  valign?: 'top' | 'middle' | 'bottom';
  
  // 边框
  border?: {
    type?: 'solid' | 'dash' | 'none';
    color?: string;
    pt?: number;
  };
  
  // 背景
  fill?: {
    color?: string;
    transparency?: number;
  };

  // 列表
  bullet?: boolean | { type?: 'number' | 'bullet' };

  // 特殊
  rowspan?: number;
  colspan?: number;
  autoPageBreak?: boolean;
  breakLine?: boolean;
}

/** 将像素值转换为 PPT 单位 */
export function PxToPPT(pxVal: string | number): number {
  const px = typeof pxVal === 'string' ? parseInt(pxVal, 10) : pxVal;
  return px / 192;
}

/** 计算元素的绝对位置（考虑父元素位置） */
export function getAbsolutePosition(element: Element): { x: number; y: number } {
  const rect = element.getBoundingClientRect();
  return {
    x: rect.left + window.scrollX,
    y: rect.top + window.scrollY
  };
}

/** 将 rgb/rgba 颜色转换为十六进制 */
export function rgbToHex(color: string): string {
  // 处理 rgba
  if (color.startsWith('rgba')) {
    const matches = color.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
    if (matches) {
      const [_, r, g, b, a] = matches;
      const hex = [r, g, b].map(x => {
        const hex = parseInt(x).toString(16);
        return hex.length === 1 ? '0' + hex : hex;
      }).join('');
      return `#${hex}`;
      // 注意：alpha 值在 PPTX 中需要单独处理为 transparency
    }
  }
  
  // 处理 rgb
  const matches = color.match(/\d+/g);
  if (!matches) return '#000000';
  
  const hex = matches.map(x => {
    const hex = parseInt(x).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
  
  return `#${hex}`;
}

/** 转换文本对齐方式 */
export function transformTextAlign(style: CSSStyleDeclaration): 'left' | 'center' | 'right' | 'justify' {
  // 优先使用 text-align
  const textAlign = style.textAlign;
  if (textAlign) {
    switch (textAlign) {
      case 'left': return 'left';
      case 'center': return 'center';
      case 'right': return 'right';
      case 'justify': return 'justify';
    }
  }
  
  // 回退到 flex 布局的对齐方式
  const justifyContent = style.justifyContent;
  switch (justifyContent) {
    case 'flex-start': return 'left';
    case 'center': return 'center';
    case 'flex-end': return 'right';
    case 'space-between':
    case 'space-around':
    case 'space-evenly':
      return 'justify';
    default:
      return 'left';
  }
}

/** 转换垂直对齐方式 */
export function transformVerticalAlign(style: CSSStyleDeclaration): 'top' | 'middle' | 'bottom' {
  // 优先使用 vertical-align
  const verticalAlign = style.verticalAlign;
  if (verticalAlign) {
    switch (verticalAlign) {
      case 'top': return 'top';
      case 'middle': return 'middle';
      case 'bottom': return 'bottom';
    }
  }
  
  // 回退到 flex 布局的对齐方式
  const alignItems = style.alignItems;
  switch (alignItems) {
    case 'flex-start': return 'top';
    case 'center': return 'middle';
    case 'flex-end': return 'bottom';
    case 'stretch': return 'middle';
    default: return 'middle';
  }
}

/** 转换字体样式 */
export function transformFontStyles(style: CSSStyleDeclaration): Pick<TransformedStyle, 'fontFace' | 'fontSize' | 'bold' | 'italic' | 'underline'> {
  const textDecoration = style.textDecoration.includes('underline');
  return {
    fontFace: style.fontFamily?.split(',')[0]?.trim().replace(/['"]/g, ''),
    fontSize: parseInt(style.fontSize, 10) * 0.3,
    bold: style.fontWeight === 'bold' || parseInt(style.fontWeight, 10) > 400,
    italic: style.fontStyle === 'italic',
    underline: textDecoration ? { style: 'single' } : undefined
  };
}

/** 转换边框样式 */
export function transformBorderStyles(style: CSSStyleDeclaration): TransformedStyle['border'] {
  const borderColor = style.borderColor;
  const borderWidth = parseInt(style.borderWidth, 10);
  const borderStyle = style.borderStyle;

  if (!borderColor || !borderWidth || borderStyle === 'none') return undefined;

  return {
    type: borderStyle === 'dashed' ? 'dash' : 'solid',
    color: rgbToHex(borderColor),
    pt: PxToPPT(borderWidth)
  };
}

/** 转换背景样式 */
export function transformBackgroundStyles(style: CSSStyleDeclaration): TransformedStyle['fill'] {
  const bgcolor = style.backgroundColor;
  if (!bgcolor || bgcolor === 'transparent') return undefined;

  let transparency = 0;
  if (bgcolor.startsWith('rgba')) {
    const matches = bgcolor.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
    if (matches) {
      transparency = Math.round((1 - parseFloat(matches[4])) * 100);
    }
  }

  return {
    color: rgbToHex(bgcolor),
    transparency
  };
}

/** 转换元素的完整样式 */
export function transformElementStyle(element: Element): TransformedStyle {
  const style = window.getComputedStyle(element);
  const box = element.getBoundingClientRect();
  const pos = getAbsolutePosition(element);

  const transformed: TransformedStyle = {
    // 位置和尺寸
    x: PxToPPT(pos.x),
    y: PxToPPT(pos.y),
    w: PxToPPT(box.width),
    h: PxToPPT(box.height),

    // 边距
    margin: [
      PxToPPT(style.marginTop),
      PxToPPT(style.marginRight),
      PxToPPT(style.marginBottom),
      PxToPPT(style.marginLeft)
    ],
    padding: [
      PxToPPT(style.paddingTop),
      PxToPPT(style.paddingRight),
      PxToPPT(style.paddingBottom),
      PxToPPT(style.paddingLeft)
    ],

    // 文本对齐
    align: transformTextAlign(style),
    valign: transformVerticalAlign(style),

    // 文本样式
    ...transformFontStyles(style),
    color: rgbToHex(style.color),

    // 边框
    border: transformBorderStyles(style),

    // 背景
    fill: transformBackgroundStyles(style),

    // 列表样式
    bullet: element.tagName.toLowerCase() === 'li'
  };

  return transformed;
}

/** 获取元素的所有可用样式属性 */
export function getElementStyles(element: Element): TransformedStyle {
  return transformElementStyle(element);
}