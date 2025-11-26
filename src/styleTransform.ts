/**
 * styleTransform.ts
 * 负责将 DOM 样式转换为 PPTXGenJS 样式对象
 */

export interface ElementStyle {
  x: number;
  y: number;
  w: number;
  h: number;
  fontSize: number;
  fontFace: string;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  align: "left" | "center" | "right" | "justify";
  valign: "top" | "middle" | "bottom";
  lineSpacing: number; // 行高 (磅)
  charSpacing: number; // 字间距 (磅)
  fill?: { color: string; transparency?: number };
  border?: any;
  opacity?: number;
  padding?: any;
}

/**
 * 颜色解析：确保返回 6 位 HEX，默认黑色
 */
export function colorToHex(color: string): string {
  if (!color || color === "transparent" || color === "inherit") return "";
  if (color.startsWith("#")) return color.replace("#", "").toUpperCase();

  if (color.startsWith("rgb")) {
    const rgba = color.match(/(\d+(\.\d+)?)/g);
    if (rgba && rgba.length >= 3) {
      // Alpha 为 0 视为透明
      if (rgba.length > 3 && parseFloat(rgba[3]) === 0) return "";
      const r = parseInt(rgba[0]).toString(16).padStart(2, "0");
      const g = parseInt(rgba[1]).toString(16).padStart(2, "0");
      const b = parseInt(rgba[2]).toString(16).padStart(2, "0");
      return (r + g + b).toUpperCase();
    }
  }
  return "000000";
}

/**
 * 获取透明度 (0-100%)
 */
function getTransparency(style: CSSStyleDeclaration): number {
  const opacity = parseFloat(style.opacity);
  if (!isNaN(opacity) && opacity < 1) {
    return (1 - opacity) * 100;
  }
  // 处理 rgba 背景色的透明度
  if (style.backgroundColor.startsWith("rgba")) {
    const rgba = style.backgroundColor.match(/(\d+(\.\d+)?)/g);
    if (rgba && rgba.length > 3) {
      return (1 - parseFloat(rgba[3])) * 100;
    }
  }
  return 0;
}

/**
 * 核心：获取计算样式
 */
export function getComputedElementStyle(
  element: Element,
  pageRect: DOMRect,
  globalScale: number,
  pageTransformScale: number
): ElementStyle {
  const style = window.getComputedStyle(element);
  const rect = element.getBoundingClientRect();

  // 1. 坐标和宽高转换 (基于渲染尺寸还原)
  const x = ((rect.left - pageRect.left) / pageTransformScale) * globalScale;
  const y = ((rect.top - pageRect.top) / pageTransformScale) * globalScale;
  const w = (rect.width / pageTransformScale) * globalScale;
  const h = (rect.height / pageTransformScale) * globalScale;

  // 2. 字体大小 (基于原始设计尺寸)
  const pxFontSize = parseFloat(style.fontSize) || 14;
  const fontSize = pxFontSize * globalScale * 72; // px -> pt

  // 3. 行高 (Line Height)
  // PPTXGenJS 的 lineSpacing 如果是数字，单位是 Points。
  // 浏览器 normal 通常约为 1.2 倍
  let lineSpacing: number;
  if (style.lineHeight === "normal") {
    lineSpacing = fontSize * 1.2;
  } else if (!isNaN(parseFloat(style.lineHeight))) {
    // 如果是纯数字 (倍数)，如 1.5
    if (/^\d+(\.\d+)?$/.test(style.lineHeight)) {
      lineSpacing = fontSize * parseFloat(style.lineHeight);
    } else {
      // 如果是 px 值
      const pxLineHeight = parseFloat(style.lineHeight);
      lineSpacing = pxLineHeight * globalScale * 72;
    }
  } else {
    lineSpacing = fontSize * 1.2;
  }

  // 4. 边框
  let border = undefined;
  const borderWidth = parseFloat(style.borderWidth);
  if (borderWidth > 0 && style.borderStyle !== "none" && style.borderColor) {
    border = {
      pt: borderWidth * globalScale * 72,
      color: colorToHex(style.borderColor),
      type: style.borderStyle === "dashed" ? "dash" : "solid",
    } as any;
  }

  // 5. 对齐
  let align: any = style.textAlign;
  if (align === "start") align = "left";
  if (align === "end") align = "right";

  let valign: any = "top";
  if (style.display === "flex") {
    if (style.alignItems === "center") valign = "middle";
    if (style.alignItems === "flex-end") valign = "bottom";
  } else {
    if (style.verticalAlign === "middle") valign = "middle";
    if (style.verticalAlign === "bottom") valign = "bottom";
  }

  // 6. 背景与透明度
  const bgColor = colorToHex(style.backgroundColor);
  const transparency = getTransparency(style);

  return {
    x,
    y,
    w,
    h,
    fontSize,
    fontFace:
      style.fontFamily?.split(",")[0].replace(/['"]/g, "").trim() || "黑体",
    color: colorToHex(style.color),
    bold: parseInt(style.fontWeight) >= 600 || style.fontWeight === "bold",
    italic: style.fontStyle === "italic",
    underline: style.textDecoration.includes("underline"),
    strike: style.textDecoration.includes("line-through"),
    align,
    valign,
    lineSpacing,
    charSpacing: parseFloat(style.letterSpacing) || 0,
    fill: bgColor
      ? {
        color: bgColor,
        transparency: transparency > 0 ? transparency : undefined,
      }
      : undefined,
    border,
    opacity: parseFloat(style.opacity),
  };
}
