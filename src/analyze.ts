import PptxGenJS from 'pptxgenjs';
import { getElementStyles, TransformedStyle } from './styleTransform';

/** 递归处理元素及其所有子元素 */
function processElement(element: Element, slide: PptxGenJS.Slide) {
  // 跳过隐藏元素
  if (element.getAttribute('hidden') !== null) return;

  const { attributes } = traverseElements(element);
  const elementStyle = getElementStyles(element);

  // 如果元素有type属性，处理当前元素
  if (attributes['type']) {
    const elementType = attributes['type'];
    if (elementType === 'text') {
      const content = element.textContent?.trim() || '';
      if (content) {
        slide.addText(content, {
          w: elementStyle.w,
          h: elementStyle.h,
          x: elementStyle.x,
          y: elementStyle.y,
          align: elementStyle.align,
          valign: elementStyle.valign,
          fontSize: elementStyle.fontSize,
          color: elementStyle.color,
          bold: elementStyle.bold,
          italic: elementStyle.italic,
          underline: elementStyle.underline,
          fontFace: elementStyle.fontFace,
          bullet: elementStyle.bullet,
          margin: elementStyle.margin
        });
      }
    } else if (elementType === 'image' && attributes['src']) {
      slide.addImage({
        path: attributes['src'],
        x: elementStyle.x,
        y: elementStyle.y,
        w: elementStyle.w,
        h: elementStyle.h
      });
    } else if (elementType === 'table') {
      try {
        const tableData: any = [];
        const rows = element.getElementsByTagName('tr');
        for (let row of rows) {
          const rowData = [];
          const cells = row.getElementsByTagName('td');
          for (let cell of cells) {
            rowData.push(cell.textContent?.trim() || '');
          }
          if (rowData.length > 0) {
            tableData.push(rowData);
          }
        }
        if (tableData.length > 0) {
          slide.addTable(tableData, {
            x: elementStyle.x,
            y: elementStyle.y,
            w: elementStyle.w,
            h: elementStyle.h,
            align: elementStyle.align,
            valign: elementStyle.valign,
            fontSize: elementStyle.fontSize,
            color: elementStyle.color,
            bold: elementStyle.bold,
            border: elementStyle.border
          });
        }
      } catch (error) {
        console.error('Error parsing table:', error);
      }
    }
  }

  // 递归处理所有子元素
  for (const child of Array.from(element.children)) {
    processElement(child, slide);
  }
}

/**
 * 将 html dom 转换为 pptx
 * @param pageClass 页面样式类名
 */
export function html2pptx(pageClass: string): PptxGenJS {
  const ppt = new PptxGenJS();
  const pageDoms = Array.from(document.querySelectorAll(pageClass)).filter(element => {
    // @ts-ignore
    return element.offsetWidth > 0 && element.offsetHeight > 0 && getComputedStyle(element).visibility !== 'hidden';
  });

  Array.from(pageDoms).forEach(dom => {
    const slide = ppt.addSlide();
    const elementStyle = getElementStyles(dom as Element);
    const { attributes } = traverseElements(dom as Element);

    // 设置背景（如果有）
    if (elementStyle.fill) {
      slide.background = {
        color: elementStyle.fill.color,
        transparency: elementStyle.fill.transparency
      };
    }

    // 递归处理所有子元素
    Array.from(dom.children).forEach(child => {
      processElement(child, slide);
    });
  });

  return ppt;
}

function traverseElements(element: Element): { attributes: any } {
  // 获取元素的属性
  const attributes: any = {};
  Array.from(element.attributes).forEach(attr => attributes[attr.name] = attr.value);
  return { attributes };
}
