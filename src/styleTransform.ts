// styleTransform.ts

import PptxGenJS from 'pptxgenjs';

export interface ElementStyle {
  x: number;
  y: number;
  w: number;
  h: number;

  // 文本相关（单位 pt）
  fontSize?: number;
  fontFace?: string;
  color?: string;
  align?: PptxGenJS.HAlign;
  valign?: PptxGenJS.VAlign;
  lineSpacing?: number; // pt
  charSpacing?: number; // pt

  // 填充/边框
  fill?: { color: string; transparency?: number };
  border?: { pt: number; color: string; type?: "solid" | "dash" | "dashed" | "dot" | "dotted" | "double" };

  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  padding?: number | [number, number, number, number];
}

export function colorToHex(color: string): string {
  if (!color || color === 'transparent' || color === 'inherit') return '';
  if (color.startsWith('#')) return color.replace('#', '').toUpperCase();
  if (color.startsWith('rgb')) {
    const rgba = color.match(/(\d+(\.\d+)?)/g);
    if (rgba && rgba.length >= 3) {
      if (rgba.length > 3 && parseFloat(rgba[3]) === 0) return '';
      const r = parseInt(rgba[0], 10).toString(16).padStart(2, '0');
      const g = parseInt(rgba[1], 10).toString(16).padStart(2, '0');
      const b = parseInt(rgba[2], 10).toString(16).padStart(2, '0');
      return (r + g + b).toUpperCase();
    }
  }
  return '000000';
}

function getTransparency(color: string): number | undefined {
  if (!color) return undefined;
  if (color.startsWith('rgba')) {
    const rgba = color.match(/(\d+(\.\d+)?)/g);
    if (rgba && rgba.length >= 4) {
      const a = parseFloat(rgba[3]);
      // pptxgenjs transparency: 0~100 (0=opaque)
      return Math.round((1 - a) * 100);
    }
  }
  return undefined;
}

/**
 * 将元素的 CSS 计算样式转换为 PPT 坐标/样式
 * 注意：坐标和尺寸 x/y/w/h 使用 globalScale/pageTransformScale 缩放；
 * 字体/行高/边框等“点数”属性仅进行单位转换 px->pt，不乘 globalScale，避免双重缩放。
 */
export function getComputedElementStyle(
  element: Element,
  pageRect: DOMRect,
  globalScale: number,
  pageTransformScale: number
): ElementStyle {
  const style = window.getComputedStyle(element);
  const rect = element.getBoundingClientRect();

  // 坐标/尺寸：px -> ppt英寸(基于 globalScale)，pptxgenjs 用英寸单位；这里维持你现有的比例策略：
  // 现有项目通常将 px*globalScale 再映射到 ppt 布局宽度（已在上层计算 globalScale）。保持不变。
  const x = ((rect.left - pageRect.left) / pageTransformScale) * globalScale;
  const y = ((rect.top - pageRect.top) / pageTransformScale) * globalScale;
  const w = (rect.width / pageTransformScale) * globalScale;
  const h = (rect.height / pageTransformScale) * globalScale;

  // 字体大小：px -> pt（仅单位转换，1px ≈ 0.75pt）
  const pxFontSize = parseFloat(style.fontSize) || 14;
  const fontSize = pxFontSize * 0.75;

  // 字体
  const fontFace = style.fontFamily?.split(',')[0].replace(/['"]/g, '');

  // 颜色
  const color = colorToHex(style.color);

  // 对齐
  const align = ((): PptxGenJS.HAlign | undefined => {
    switch (style.textAlign) {
      case 'left':
      case 'start':
        return 'left';
      case 'center':
        return 'center';
      case 'right':
      case 'end':
        return 'right';
      case 'justify':
        return 'justify';
      default:
        return undefined;
    }
  })();

  const valign = ((): PptxGenJS.VAlign | undefined => {
    switch (style.verticalAlign) {
      case 'top':
        return 'top';
      case 'middle':
      case 'center':
        return 'middle';
      case 'bottom':
        return 'bottom';
      default:
        return undefined;
    }
  })();

  // 行高：CSS -> pt
  let lineSpacing: number | undefined;
  const lh = style.lineHeight;
  if (lh === 'normal' || !lh) {
    lineSpacing = fontSize * 1.2;
  } else if (/^\d+(\.\d+)?$/.test(lh)) {
    // 数字倍数
    lineSpacing = fontSize * parseFloat(lh);
  } else {
    // px 或其它单位，粗略按 px -> pt
    const pxLineHeight = parseFloat(lh);
    if (!isNaN(pxLineHeight)) {
      lineSpacing = pxLineHeight * 0.75;
    }
  }

  // 字间距（letter-spacing）px->pt
  let charSpacing: number | undefined;
  const ls = style.letterSpacing;
  if (ls && ls !== 'normal') {
    const pxLs = parseFloat(ls);
    if (!isNaN(pxLs)) {
      charSpacing = pxLs * 0.75;
    }
  }

  // 背景填充
  const bgColor = colorToHex(style.backgroundColor);
  const fill =
    bgColor && bgColor !== ''
      ? { color: bgColor, transparency: getTransparency(style.backgroundColor) }
      : undefined;

  // 边框（统一采用 px->pt，不叠加 globalScale，避免双重缩放）
  let border: ElementStyle['border'];
  const borderWidth = parseFloat(style.borderWidth);
  if (borderWidth > 0 && style.borderStyle !== 'none' && style.borderColor) {
    border = {
      pt: borderWidth * 0.75,
      color: colorToHex(style.borderColor),
      type: style.borderStyle === 'dashed' ? 'dash' : 'solid',
    };
  }

  // 文本样式
  const bold = parseInt(style.fontWeight || '400', 10) >= 600 || style.fontWeight === 'bold';
  const italic = style.fontStyle === 'italic';
  const underline = style.textDecoration.includes('underline');
  const strike = style.textDecoration.includes('line-through');

  // 内边距（用于表格单元格等）
  let padding: ElementStyle['padding'];
  const pt = (v: string) => (parseFloat(v || '0') || 0) * 0.75;
  const ptTop = pt(style.paddingTop);
  const ptRight = pt(style.paddingRight);
  const ptBottom = pt(style.paddingBottom);
  const ptLeft = pt(style.paddingLeft);
  if (ptTop || ptRight || ptBottom || ptLeft) {
    padding = [ptTop, ptRight, ptBottom, ptLeft];
  }

  return {
    x,
    y,
    w,
    h,
    fontSize,
    fontFace,
    color,
    align,
    valign,
    lineSpacing,
    charSpacing,
    fill,
    border,
    bold,
    italic,
    underline,
    strike,
    padding,
  };
}