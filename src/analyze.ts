import PptxGenJS from "pptxgenjs";
import { getComputedElementStyle, colorToHex } from "./styleTransform";
import { getElementAnimation } from "./animationTransform";

const PPT_LAYOUT = {
  width: 10,
  height: 5.625,
};

interface ParsedTextShape {
  text: string;
  options: PptxGenJS.TextPropsOptions;
}

interface ParseResult {
  shapes: ParsedTextShape[]; // 以前是 textItems，现在改为 shapes
  consumedElements: Set<Element>;
}

// 新增：图表配置接口，对应 pptxgenjs 的 addChart 参数
interface PptChartConfig {
  type: any;
  data: PptxGenJS.OptsChartData[];
  options?: PptxGenJS.IChartOpts;
}

/**
 * 辅助函数：生成模拟间距的空格字符串
 */
function getSpaceString(widthPx: number, fontSizePx: number): string {
  if (widthPx <= 0) return "";
  const spaceWidth = fontSizePx / 3;
  const count = Math.round(widthPx / spaceWidth);
  return " ".repeat(count);
}

/**
 * 解析富文本 - 独立坐标模式
 * 不再聚合文本，而是计算每个 TextNode 的真实屏幕位置
 */
function parseRichText(
  rootElement: Element,
  pageRect: DOMRect, // 需要传入页面容器的 rect 用于计算相对坐标
  globalScale: number,
  pageTransformScale: number
): ParseResult {
  const shapes: ParsedTextShape[] = [];
  const consumedElements = new Set<Element>();

  function traverse(node: Node, parentStyle: CSSStyleDeclaration) {
    // 1. 处理文本节点
    if (node.nodeType === Node.TEXT_NODE) {
      const textContent = node.textContent || "";
      if (!textContent.trim()) return;

      const range = document.createRange();
      range.selectNode(node);
      const rect = range.getBoundingClientRect();

      if (rect.width === 0 || rect.height === 0) return;

      // --- 坐标计算 (逻辑同 getComputedElementStyle) ---
      // 计算相对于 PPT 页面的坐标
      const x = ((rect.left - pageRect.left) / pageTransformScale) * globalScale;
      const y = ((rect.top - pageRect.top) / pageTransformScale) * globalScale;

      // 宽度缓冲：为了防止 PPT 换行，稍微给宽一点点 (例如 + 2px 对应的 ppt单位)
      // 也可以选择不加缓冲，但 PPT 渲染引擎通常比浏览器占宽
      const w = (rect.width / pageTransformScale) * globalScale + 0.05;
      const h = (rect.height / pageTransformScale) * globalScale;

      // --- 样式计算 ---
      const pxSize = parseFloat(parentStyle.fontSize || "14");
      const ptSize = pxSize * globalScale * 72;
      const bgColor = colorToHex(parentStyle.backgroundColor);

      // 将该文本节点作为一个独立的文本框对象存入
      shapes.push({
        text: textContent,
        options: {
          x,
          y,
          w,
          h,
          fontSize: ptSize,
          color: colorToHex(parentStyle.color),
          fontFace: parentStyle.fontFamily?.split(",")[0].replace(/['"]/g, ""),
          bold: parseInt(parentStyle.fontWeight) >= 600 || parentStyle.fontWeight === "bold",
          italic: parentStyle.fontStyle === "italic",
          underline: parentStyle.textDecoration.includes("underline") ? { style: "sng" } : undefined,
          strike: parentStyle.textDecoration.includes("line-through"),
          highlight: bgColor ? bgColor : undefined,
          subscript: parentStyle.verticalAlign === "sub",
          superscript: parentStyle.verticalAlign === "super",
          align: "left", // 强制左对齐，因为 x 坐标已经是文本的起点了
          valign: "top", // 强制顶对齐
          autoFit: false,
          wrap: true, // 允许换行，防止溢出，但通常 rect 已经是一行的宽度了
          inset: 0, // 去除内边距，保证位置精准
        },
      });
    }
    // 2. 处理元素节点
    else if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;

      const stopTags = ["TABLE", "IMG", "CANVAS", "SVG", "VIDEO", "IFRAME"];
      if (stopTags.includes(el.tagName)) {
        // 直接 return，不加入 consumedElements。
        // 这样 processElement 的递归循环 (Array.from(element.children))
        // 依然会遍历到这个节点，并调用专门的 addTable/addImage 逻辑。
        return;
      }
      const style = window.getComputedStyle(el);
      if (el.tagName === "BR") {
        // BR 标签在独立坐标模式下通常不需要处理，
        // 因为下一个文本节点的 Range rect 会自动出现在下一行
        consumedElements.add(el);
        return;
      }

      // 布局边界检查 (同上一次修改)
      const isBlock = style.display === "block" || style.display === "flex" || style.display === "grid" || style.display === "inline-block" || style.position === "absolute" || style.position === "fixed";

      const isStyleTag = ["SPAN", "B", "STRONG", "I", "EM", "U", "FONT", "SUB", "SUP", "A", "SMALL", "BIG"].includes(el.tagName);

      // 如果是块级元素且不是样式标签，跳过，留给 processElement
      if (isBlock && !isStyleTag) {
        return;
      }

      if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") {
        consumedElements.add(el);
        return;
      }

      consumedElements.add(el);
      el.childNodes.forEach((child) => traverse(child, style));
    }
  }

  const rootStyle = window.getComputedStyle(rootElement);
  rootElement.childNodes.forEach((child) => traverse(child, rootStyle));

  return { shapes, consumedElements };
}
/**
 * 处理单个元素
 */
function processElement(element: Element, slide: PptxGenJS.Slide | any, pageRect: DOMRect, globalScale: number, pageTransformScale: number) {
  if (element.getAttribute("hidden") !== null) return;
  const style = window.getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") return;
  const pptStyle = getComputedElementStyle(element, pageRect, globalScale, pageTransformScale);
  const animation = getElementAnimation(element);

  // ============================================================
  // --- 0. 优先处理原生图表 (Native Chart) ---
  // 检查是否存在 data-pptx-chart-config 属性
  // ============================================================
  const chartConfigStr = element.getAttribute("data-pptx-chart-config");
  if (chartConfigStr) {
    try {
      // 解析配置
      const chartConfig: PptChartConfig = JSON.parse(chartConfigStr);

      // 合并样式：DOM 的位置 + JSON 中的配置
      // JSON 中的 options 优先级更高，允许用户覆盖自动计算的 x,y,w,h
      const finalOptions: PptxGenJS.IChartOpts = {
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h,
        ...chartConfig.options,
      };

      // 添加原生图表
      slide.addChart(chartConfig.type, chartConfig.data, finalOptions);

      // 如果成功处理了图表，直接返回，不再作为图片或文本处理
      return;
    } catch (e) {
      console.warn("解析图表配置失败，将回退到截图模式:", e);
      // 解析失败不 return，继续往下走，尝试作为 Canvas/Image 截图处理
    }
  }

  // --- 1. 处理图片 ---
  if (element.tagName === "IMG") {
    const imgEl = element as HTMLImageElement;
    if (imgEl.src) {
      slide.addImage({
        path: imgEl.src,
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h,
        sizing: { type: "contain", w: pptStyle.w, h: pptStyle.h },
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {}),
      });
    }
    return;
  }

  // --- 2. 处理 Canvas (图表的回退方案) ---
  if (element.tagName === "CANVAS") {
    try {
      const canvas = element as HTMLCanvasElement;
      // 增加判断：如果是空 canvas 或者 tainted canvas 可能会报错
      const imgData = canvas.toDataURL("image/png");
      slide.addImage({
        data: imgData,
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h,
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {}),
      });
    } catch (e) {
      console.warn("Canvas export failed", e);
    }
    return;
  }

  // --- 3. 处理表格 ---
  if (element.tagName === "TABLE") {
    // ... (保持原有的 Table 处理逻辑不变) ...
    const tableElement = element as HTMLTableElement;
    const rows = Array.from(tableElement.querySelectorAll("tr"));
    if (rows.length === 0) return;

    const colWidthsPx: number[] = [];
    let maxCols = 0;
    rows.forEach((row) => {
      let currentCols = 0;
      Array.from(row.children).forEach((cell) => {
        currentCols += parseInt(cell.getAttribute("colspan") || "1");
      });
      if (currentCols > maxCols) maxCols = currentCols;
    });

    for (let i = 0; i < maxCols; i++) colWidthsPx.push(0);

    rows.forEach((row) => {
      let currentColIdx = 0;
      Array.from(row.children).forEach((cell) => {
        const cellRect = cell.getBoundingClientRect();
        const colSpan = parseInt(cell.getAttribute("colspan") || "1");
        const cellPxWidth = cellRect.width;
        const avgColWidth = cellPxWidth / colSpan;
        for (let i = 0; i < colSpan; i++) {
          if (colWidthsPx[currentColIdx + i] === 0 || avgColWidth > colWidthsPx[currentColIdx + i]) {
            colWidthsPx[currentColIdx + i] = avgColWidth;
          }
        }
        currentColIdx += colSpan;
      });
    });

    const colW: number[] = colWidthsPx.map((px) => (px / pageTransformScale) * globalScale);

    const totalPptColWidth = colW.reduce((sum, val) => sum + val, 0);
    if (totalPptColWidth > 0 && Math.abs(totalPptColWidth - pptStyle.w) > 0.01) {
      const scaleFactor = pptStyle.w / totalPptColWidth;
      for (let i = 0; i < colW.length; i++) colW[i] *= scaleFactor;
    } else if (colW.length === 0 && rows.length > 0) {
      colW.push(...Array(maxCols > 0 ? maxCols : 1).fill(pptStyle.w / (maxCols > 0 ? maxCols : 1)));
    }

    const rowH: number[] = [];
    rows.forEach((row) => {
      const rowRect = row.getBoundingClientRect();
      rowH.push((rowRect.height / pageTransformScale) * globalScale);
    });

    const tableData: PptxGenJS.TableRow[] = [];
    rows.forEach((row) => {
      const rowData: PptxGenJS.TableCell[] | any = [];
      const cells = Array.from(row.querySelectorAll("td, th"));

      cells.forEach((cell) => {
        const cellStyle = getComputedElementStyle(cell, pageRect, globalScale, pageTransformScale);
        const cellTxt = cell.textContent || "";

        rowData.push({
          text: cellTxt,
          options: {
            fill: cellStyle.fill,
            color: cellStyle.color,
            bold: cellStyle.bold,
            italic: cellStyle.italic,
            underline: cellStyle.underline,
            strike: cellStyle.strike,
            align: cellStyle.align,
            valign: cellStyle.valign,
            margin: cellStyle.padding,
            border: cellStyle.border
              ? {
                  pt: cellStyle.border.pt,
                  color: cellStyle.border.color,
                  type: cellStyle.border.type as any,
                }
              : undefined,
            rowspan: parseInt(cell.getAttribute("rowspan") || "1"),
            colspan: parseInt(cell.getAttribute("colspan") || "1"),
            fontSize: cellStyle.fontSize,
            fontFace: cellStyle.fontFace,
            lineSpacing: cellStyle.lineSpacing,
            charSpacing: cellStyle.charSpacing,
            wrap: true,
            autoFit: false,
          },
        });
      });
      if (rowData.length) tableData.push(rowData);
    });

    if (tableData.length) {
      slide.addTable(tableData, {
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        colW: colW.length > 0 ? colW : undefined,
        rowH: rowH.length > 0 ? rowH : undefined,
        fill: pptStyle.fill,
        line: pptStyle.border,
      });
    }
    return;
  }
  // --- 4. 预先计算文本 (先不画，只计算) ---
  const { shapes, consumedElements } = parseRichText(element, pageRect, globalScale, pageTransformScale);

  // --- 5. 普通容器 (背景/边框/动画) ---
  const hasBackground = pptStyle.fill && pptStyle.fill.color;
  const hasBorder = pptStyle.border;
  if (hasBackground || hasBorder || (shapes.length === 0 && !["SPAN", "B", "I", "U", "STRONG", "EM"].includes(element.tagName))) {
    slide.addShape(pptStyle.shapeType, {
      x: pptStyle.x,
      y: pptStyle.y,
      w: pptStyle.w,
      h: pptStyle.h,
      fill: pptStyle.fill,
      line: pptStyle.border,
      rectRadius: pptStyle.rectRadius,
      rotate: pptStyle.rotate,
      ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {}),
    });
  }
  // --- 6. 绘制文本 ---
  if (shapes.length > 0) {
    shapes.forEach((shape) => {
      slide.addText(shape.text, {
        ...shape.options,
        // 继承父级动画
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {}),
        // 继承父级旋转
        rotate: pptStyle.rotate,
      });
    });
  }

  // --- 递归处理子元素 ---
  Array.from(element.children).forEach((child) => {
    if (consumedElements.has(child)) return;
    processElement(child, slide, pageRect, globalScale, pageTransformScale);
  });
}

/**
 * 将 html dom 转换为 pptx 对象
 */
export function html2pptx(pageClass: string): PptxGenJS {
  const ppt = new PptxGenJS();
  ppt.layout = "LAYOUT_16x9";
  ppt.theme = { headFontFace: "黑体" };
  const pages = document.querySelectorAll(`.${pageClass}`);

  if (pages.length === 0) {
    console.warn(`[html2pptx] 未找到类名为 .${pageClass} 的幻灯片元素`);
    return ppt;
  }

  pages.forEach((dom) => {
    const element = dom as HTMLElement;
    if (element.offsetWidth === 0 || element.offsetHeight === 0) return;

    const pageRect = element.getBoundingClientRect();
    const pageTransformScale = pageRect.width / element.offsetWidth;

    if (pageTransformScale === 0) return;

    const unscaledPageWidth = element.offsetWidth;
    const globalScale = PPT_LAYOUT.width / unscaledPageWidth;

    const slide = ppt.addSlide();

    const bgStyle = window.getComputedStyle(element);
    const bgImg = bgStyle.backgroundImage;
    const bgColor = colorToHex(bgStyle.backgroundColor);
    if (bgImg && bgImg.includes("url(")) {
      const match = bgImg.match(/url\s*\(['"]?(.*?)['"]?\)/);
      const imageUrl = match && match[1] ? match[1] : undefined;
      slide.addImage({ x: 0, y: 0, w: "100%", h: "100%", path: imageUrl });
    } else if (bgColor) {
      // 处理纯色 (如果背景是渐变，通常 computedStyle 的 backgroundColor 会是透明或回退色，这里只取 color)
      slide.background = { color: bgColor };
    }
    // console.log("page:", slide);
    Array.from(element.children).forEach((child) => {
      processElement(child, slide, pageRect, globalScale, pageTransformScale);
    });
  });

  return ppt;
}

/**
 * 从 URL 获取图片并将其编码为 Base64 Data URI。
 * @param url 图片的 URL
 * @returns Base64 编码的图片字符串，如果失败则返回 null。
 */
async function fetchAndEncodeImageToBase64(url: string): Promise<string | null> {
  try {
    const response = await fetch(url, { mode: "cors" }); // 尝试使用 cors 模式
    if (!response.ok) {
      console.error(`Failed to fetch image from ${url}: ${response.statusText}`);
      return null;
    }
    const blob = await response.blob(); // 获取图片 Blob 数据

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob); // 将 Blob 转换为 Data URL (Base64)
    });
  } catch (error) {
    console.error(`Error fetching or encoding image from ${url}:`, error);
    return null;
  }
}
