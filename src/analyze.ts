import PptxGenJS from 'pptxgenjs';
import { getComputedElementStyle, colorToHex } from './styleTransform';
import { getElementAnimation } from './animationTransform';

const PPT_LAYOUT = {
  width: 10,
  height: 5.625
};

interface ParseResult {
  textItems: any[];
  consumedElements: Set<Element>;
}

// 新增：图表配置接口，对应 pptxgenjs 的 addChart 参数
interface PptChartConfig {
  type: any;
  data: any[];
  options?: PptxGenJS.IChartOpts;
}

/**
 * 辅助函数：生成模拟间距的空格字符串
 */
function getSpaceString(widthPx: number, fontSizePx: number): string {
  if (widthPx <= 0) return '';
  const spaceWidth = fontSizePx / 3;
  const count = Math.round(widthPx / spaceWidth);
  return ' '.repeat(count);
}

/**
 * 解析富文本
 */
function parseRichText(rootElement: Element, globalScale: number, pageTransformScale: number): ParseResult {
  // ... (保持原有的 parseRichText 逻辑不变，这里省略以节省篇幅) ...
  const textItems: any[] = [];
  const consumedElements = new Set<Element>();
  let lastRightPos: number | null = null;

  function traverse(node: Node, parentStyle: CSSStyleDeclaration) {
    const currentFontSizePx = parseFloat(parentStyle.fontSize || '14');

    if (node.nodeType === Node.TEXT_NODE) {
      const whiteSpace = parentStyle.whiteSpace;
      let textContent = node.textContent || '';
      if (whiteSpace !== 'pre' && whiteSpace !== 'pre-wrap') {
        textContent = textContent.replace(/[\n\r]+/g, '');
      }
      if (!textContent && textContent.indexOf('\n') === -1) return;
      if (textContent.trim() === '' && whiteSpace !== 'pre') return;

      const parentEl = node.parentElement;
      if (parentEl) {
        const rect = parentEl.getBoundingClientRect();
        if (lastRightPos !== null) {
          const gapPx = rect.left - lastRightPos;
          if (gapPx > 2) {
            const spaces = getSpaceString(gapPx / pageTransformScale, currentFontSizePx);
            if (spaces) textItems.push({ text: spaces });
          }
        }
        lastRightPos = rect.right;
      }

      const pxSize = parseFloat(parentStyle.fontSize || '14');
      const ptSize = pxSize * globalScale * 72;
      const bgColor = colorToHex(parentStyle.backgroundColor);

      textItems.push({
        text: textContent,
        options: {
          color: colorToHex(parentStyle.color),
          fontSize: ptSize,
          bold: parseInt(parentStyle.fontWeight) >= 600 || parentStyle.fontWeight === 'bold',
          italic: parentStyle.fontStyle === 'italic',
          underline: parentStyle.textDecoration.includes('underline') ? { style: 'sng' } : undefined,
          strike: parentStyle.textDecoration.includes('line-through'),
          fontFace: parentStyle.fontFamily?.split(',')[0].replace(/['"]/g, ''),
          subscript: parentStyle.verticalAlign === 'sub',
          superscript: parentStyle.verticalAlign === 'super',
          highlight: bgColor ? bgColor : undefined
        }
      });
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;
      const style = window.getComputedStyle(el);

      if (el.tagName === 'BR') {
        textItems.push({ text: '', options: { breakLine: true } });
        consumedElements.add(el);
        lastRightPos = null;
        return;
      }

      const isPureInline = style.display === 'inline';
      const isStyleTag = ['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'FONT', 'SUB', 'SUP', 'A'].includes(el.tagName);

      if ((!isPureInline && !isStyleTag) || ['IMG', 'CANVAS', 'TABLE', 'SVG'].includes(el.tagName)) {
        lastRightPos = null;
        return;
      }

      if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') {
        consumedElements.add(el);
        return;
      }

      consumedElements.add(el);
      el.childNodes.forEach(child => traverse(child, style));
    }
  }

  const rootStyle = window.getComputedStyle(rootElement);
  rootElement.childNodes.forEach(child => traverse(child, rootStyle));

  return { textItems, consumedElements };
}

/**
 * 处理单个元素
 */
function processElement(element: Element, slide: any, pageRect: DOMRect, globalScale: number, pageTransformScale: number) {
  if (element.getAttribute('hidden') !== null) return;
  const style = window.getComputedStyle(element);
  if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return;

  const pptStyle = getComputedElementStyle(element, pageRect, globalScale, pageTransformScale);
  const animation = getElementAnimation(element);

  // ============================================================
  // --- 0. 优先处理原生图表 (Native Chart) ---
  // 检查是否存在 data-pptx-chart-config 属性
  // ============================================================
  const chartConfigStr = element.getAttribute('data-pptx-chart-config');
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
        ...chartConfig.options
      };

      // 添加原生图表
      slide.addChart(chartConfig.type, chartConfig.data, finalOptions);

      // 如果成功处理了图表，直接返回，不再作为图片或文本处理
      return;
    } catch (e) {
      console.warn('解析图表配置失败，将回退到截图模式:', e);
      // 解析失败不 return，继续往下走，尝试作为 Canvas/Image 截图处理
    }
  }

  // --- 1. 处理图片 ---
  if (element.tagName === 'IMG') {
    const imgEl = element as HTMLImageElement;
    if (imgEl.src) {
      slide.addImage({
        path: imgEl.src,
        x: pptStyle.x, y: pptStyle.y, w: pptStyle.w, h: pptStyle.h,
        sizing: { type: 'contain', w: pptStyle.w, h: pptStyle.h },
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {})
      });
    }
    return;
  }

  // --- 2. 处理 Canvas (图表的回退方案) ---
  if (element.tagName === 'CANVAS') {
    try {
      const canvas = element as HTMLCanvasElement;
      // 增加判断：如果是空 canvas 或者 tainted canvas 可能会报错
      const imgData = canvas.toDataURL('image/png');
      slide.addImage({
        data: imgData,
        x: pptStyle.x, y: pptStyle.y, w: pptStyle.w, h: pptStyle.h,
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {})
      });
    } catch (e) {
      console.warn('Canvas export failed', e);
    }
    return;
  }

  // --- 3. 处理表格 ---
  if (element.tagName === 'TABLE') {
    // ... (保持原有的 Table 处理逻辑不变) ...
    const tableElement = element as HTMLTableElement;
    const rows = Array.from(tableElement.querySelectorAll('tr'));
    if (rows.length === 0) return;

    const colWidthsPx: number[] = [];
    let maxCols = 0;
    rows.forEach(row => {
      let currentCols = 0;
      Array.from(row.children).forEach(cell => {
        currentCols += parseInt(cell.getAttribute('colspan') || '1');
      });
      if (currentCols > maxCols) maxCols = currentCols;
    });

    for (let i = 0; i < maxCols; i++) colWidthsPx.push(0);

    rows.forEach(row => {
      let currentColIdx = 0;
      Array.from(row.children).forEach(cell => {
        const cellRect = cell.getBoundingClientRect();
        const colSpan = parseInt(cell.getAttribute('colspan') || '1');
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

    const colW: number[] = colWidthsPx.map(px => (px / pageTransformScale) * globalScale);

    const totalPptColWidth = colW.reduce((sum, val) => sum + val, 0);
    if (totalPptColWidth > 0 && Math.abs(totalPptColWidth - pptStyle.w) > 0.01) {
      const scaleFactor = pptStyle.w / totalPptColWidth;
      for (let i = 0; i < colW.length; i++) colW[i] *= scaleFactor;
    } else if (colW.length === 0 && rows.length > 0) {
      colW.push(...Array(maxCols > 0 ? maxCols : 1).fill(pptStyle.w / (maxCols > 0 ? maxCols : 1)));
    }

    const rowH: number[] = [];
    rows.forEach(row => {
      const rowRect = row.getBoundingClientRect();
      rowH.push((rowRect.height / pageTransformScale) * globalScale);
    });

    const tableData: PptxGenJS.TableRow[] = [];
    rows.forEach(row => {
      const rowData: any[] = [];
      const cells = Array.from(row.querySelectorAll('td, th'));

      cells.forEach(cell => {
        const cellStyle = getComputedElementStyle(cell, pageRect, globalScale, pageTransformScale);
        const cellTxt = cell.textContent || '';

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
            border: cellStyle.border ? {
              pt: cellStyle.border.pt,
              color: cellStyle.border.color,
              type: cellStyle.border.type as any
            } : undefined,
            rowspan: parseInt(cell.getAttribute('rowspan') || '1'),
            colspan: parseInt(cell.getAttribute('colspan') || '1'),
            fontSize: cellStyle.fontSize,
            fontFace: cellStyle.fontFace,
            lineSpacing: cellStyle.lineSpacing,
            charSpacing: cellStyle.charSpacing,
            wrap: true,
            autoFit: false
          }
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
        line: pptStyle.border
      });
    }
    return;
  }

  // --- 4. 处理文本 ---
  const { textItems, consumedElements } = parseRichText(element, globalScale, pageTransformScale);

  if (textItems.length > 0) {
    const pxToInch = (pxStr: string) => (parseFloat(pxStr) || 0) * globalScale;
    const inset: [number, number, number, number] = [
      pxToInch(style.paddingTop),
      pxToInch(style.paddingRight),
      pxToInch(style.paddingBottom),
      pxToInch(style.paddingLeft)
    ];

    slide.addText(textItems, {
      x: pptStyle.x,
      y: pptStyle.y,
      w: pptStyle.w,
      h: pptStyle.h,
      align: pptStyle.align,
      valign: pptStyle.valign,
      fill: pptStyle.fill,
      line: pptStyle.border,
      lineSpacing: pptStyle.lineSpacing,
      charSpacing: pptStyle.charSpacing,
      inset: inset,
      wrap: true,
      autoFit: false,
      isTextBox: true,
      ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {})
    });
  } else {
    // --- 5. 普通容器 (背景/边框/动画) ---
    if ((pptStyle.fill && pptStyle.fill.color) || pptStyle.border || animation) {
      slide.addShape('rect', {
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h,
        fill: pptStyle.fill,
        line: pptStyle.border,
        ...(animation ? { animate: { type: animation.type, duration: animation.duration } } : {})
      });
    }
  }

  // --- 6. 递归处理子元素 ---
  Array.from(element.children).forEach(child => {
    if (consumedElements.has(child)) return;
    processElement(child, slide, pageRect, globalScale, pageTransformScale);
  });
}

/**
 * 将 html dom 转换为 pptx 对象
 */
export function html2pptx(pageClass: string): PptxGenJS {
  // ... (保持原有的 html2pptx 逻辑不变) ...
  const ppt = new PptxGenJS();
  ppt.layout = 'LAYOUT_16x9';

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
    const bgColor = colorToHex(bgStyle.backgroundColor);
    if (bgColor) {
      slide.background = { color: bgColor };
    }

    Array.from(element.children).forEach(child => {
      processElement(child, slide, pageRect, globalScale, pageTransformScale);
    });
  });

  return ppt;
}