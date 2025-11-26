/**
 * analyze.ts
 */
import PptxGenJS from 'pptxgenjs';
import { getComputedElementStyle, colorToHex } from './styleTransform';

const PPT_LAYOUT = {
  width: 10,
  height: 5.625
};

function getElementTransformScale(element: Element): number {
  const style = window.getComputedStyle(element);
  const transform = style.transform;
  if (transform && transform !== 'none') {
    const matrix = transform.match(/matrix\(([^)]+)\)/);
    if (matrix && matrix[1]) {
      const values = matrix[1].split(',').map(Number);
      return values[0];
    }
  }
  return 1;
}

/**
 * 解析富文本
 * 修复：增加了对 TABLE, CANVAS, IMG 等独立渲染元素的阻断，防止父级容器提取其内部文本
 */
function parseRichText(element: Element, globalScale: number): any[] {
  const textItems: any[] = [];

  function traverse(node: Node, parentStyle: CSSStyleDeclaration) {
    if (node.nodeType === Node.TEXT_NODE) {
      const whiteSpace = parentStyle.whiteSpace;
      let textContent = node.textContent || '';
      if (whiteSpace === 'pre' || whiteSpace === 'pre-wrap') {
        // 保留换行符
      } else {
        // 移除换行符，依靠 wrap: true
        textContent = textContent.replace(/[\n\r]+/g, '');
      }

      if (!textContent.trim() && textContent.indexOf('\n') === -1) return;

      const pxSize = parseFloat(parentStyle.fontSize || '14');
      const ptSize = pxSize * globalScale * 72;

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
        }
      });
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;

      // --- 关键修复开始 ---
      // 如果遇到这些标签，说明它们是独立的组件（表格、图表、图片），
      // 它们的文本内容属于它们自己，不应该被父级容器提取。
      // 直接 return，不再递归遍历其子节点。
      if (['TABLE', 'CANVAS', 'IMG', 'SVG', 'VIDEO', 'AUDIO'].includes(el.tagName)) {
        return;
      }
      // --- 关键修复结束 ---

      const style = window.getComputedStyle(el);

      if (el.tagName === 'BR') {
        textItems.push({ text: '', options: { breakLine: true } });
        return;
      }
      if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return;

      el.childNodes.forEach(child => traverse(child, style));
    }
  }

  traverse(element, window.getComputedStyle(element));
  return textItems;
}

/**
 * 处理单个元素
 * (此处代码逻辑保持不变，使用上一次修复了表格行高和文本堆叠的版本)
 */
function processElement(element: Element, slide: PptxGenJS.Slide, pageRect: DOMRect, globalScale: number, pageTransformScale: number, processedTextParent: boolean = false) {
  // 基础过滤
  if (element.getAttribute('hidden') !== null) return;
  const style = window.getComputedStyle(element);
  if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return;

  const pptStyle = getComputedElementStyle(element, pageRect, globalScale, pageTransformScale);

  // --- 1. 处理图片 (IMG 标签) ---
  if (element.tagName === 'IMG') {
    const imgEl = element as HTMLImageElement;
    const src = imgEl.src;
    if (src) {
      slide.addImage({
        path: src,
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h,
        sizing: { type: 'contain', w: pptStyle.w, h: pptStyle.h }
      });
    }
    return; // 图片元素是原子性的，不递归处理其子元素
  }

  // --- 2. 处理 Canvas (通常是图表) ---
  if (element.tagName === 'CANVAS') {
    const canvas = element as HTMLCanvasElement;
    try {
      const originalBackgroundColor = canvas.style.backgroundColor;
      canvas.style.backgroundColor = 'white';
      const imgData = canvas.toDataURL('image/png', 1.0);
      canvas.style.backgroundColor = originalBackgroundColor;

      slide.addImage({
        data: imgData,
        x: pptStyle.x,
        y: pptStyle.y,
        w: pptStyle.w,
        h: pptStyle.h
      });
    } catch (e) {
      console.warn('Canvas export failed', e);
    }
    return;
  }

  // --- 3. 处理表格 (TABLE 标签) ---
  if (element.tagName === 'TABLE') {
    const tableElement = element as HTMLTableElement;
    const rows = Array.from(tableElement.querySelectorAll('tr'));
    if (rows.length === 0) return;

    // 3.1 计算列宽
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

    // 3.2 计算每行的实际高度
    const rowH: number[] = [];
    rows.forEach(row => {
      const rowRect = row.getBoundingClientRect();
      rowH.push((rowRect.height / pageTransformScale) * globalScale);
    });

    // 3.3 构建数据
    const tableData: PptxGenJS.TableRow[] = [];
    rows.forEach(row => {
      // const rowData: PptxGenJS.TableCell[] = [];
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
      } as any);
    }
    return;
  }

  // --- 4. 处理文本 ---
  let currentElementProcessedAsText = false;
  if (!processedTextParent) {
    const textItems = parseRichText(element, globalScale);
    if (textItems.length > 0) {
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
        isTextBox: true,
        wrap: true,
        autoFit: false,
      });
      currentElementProcessedAsText = true;
    }
  }

  // --- 5. 普通容器 (背景/边框) ---
  if (!currentElementProcessedAsText && ((pptStyle.fill && pptStyle.fill.color) || pptStyle.border)) {
    slide.addShape('rect', {
      x: pptStyle.x,
      y: pptStyle.y,
      w: pptStyle.w,
      h: pptStyle.h,
      fill: pptStyle.fill,
      line: pptStyle.border
    });
  }

  // --- 6. 递归处理子元素 ---
  Array.from(element.children).forEach(child => {
    processElement(child, slide, pageRect, globalScale, pageTransformScale, currentElementProcessedAsText || processedTextParent);
  });
}

/**
 * 将 html dom 转换为 pptx 对象
 */
export function html2pptx(pageClass: string): PptxGenJS {
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

    const pageTransformScale = getElementTransformScale(element);
    if (pageTransformScale === 0) return;

    const pageRect = element.getBoundingClientRect();
    const unscaledPageWidth = pageRect.width / pageTransformScale;

    const globalScale = PPT_LAYOUT.width / unscaledPageWidth;

    const slide = ppt.addSlide();

    const bgStyle = window.getComputedStyle(element);
    const bgColor = colorToHex(bgStyle.backgroundColor);
    if (bgColor) {
      slide.background = { color: bgColor };
    }

    Array.from(element.children).forEach(child => {
      processElement(child, slide, pageRect, globalScale, pageTransformScale, false);
    });
  });

  return ppt;
}
