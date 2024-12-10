import PptxGenJS from 'pptxgenjs';


/**
 * 将 markdown 转换为 pptx
 * @param pageClass 页面样式类名
 * @param ppt pptx 实例对象
 */
export function markdown2pptx(pageClass: string): PptxGenJS {
  const ppt = new PptxGenJS();
  const pageDoms = Array.from(document.querySelectorAll(pageClass)).filter(element => {
    // @ts-ignore
    return element.offsetWidth > 0 && element.offsetHeight > 0 && getComputedStyle(element).visibility !== 'hidden';
  });

  Array.from(pageDoms).forEach(dom => {
    const slide = ppt.addSlide();
    const {styles, attributes} = traverseElements(dom as Element); // 调用遍历函数
    const backgroundImage = styles['background-image'] && styles['background-image']?.match(/url\(["']?([^"']*)["']?\)/)?.length ? styles['background-image']?.match(/url\(["']?([^"']*)["']?\)/)[1] : null;
    slide.background = {color: rgbToHex(styles['background-color']), path: backgroundImage};
    // 遍历子元素
    dom.childNodes.forEach((child: any) => {
      if (child.nodeType === Node.ELEMENT_NODE) { // 确保是元素节点
        const text = child.innerHTML;
        const {styles, attributes} = traverseElements(child as Element); // 调用遍历函数
        const {width, height, top, left, color} = styles;
        const alignItems = styles['align-items'];
        const justifyContent = styles['justify-content'];
        const fontSize = parseInt(styles['font-size'], 10) * .3;
        const fontWeight = styles['font-weight'];
        const borderColor = styles['border-color'];
        if (attributes['type'] === 'text') {
          if (text.startsWith('<ul') || text.startsWith('<li')) {
            let listItems;
            listItems = child.querySelectorAll('li');


            let computeHeight = 0;
            listItems.forEach((item: Element) => {
              const {styles: liStyles} = traverseElements(item); // 调用遍历函数
              const {
                width: liWidth,
                height: liHeight,
                top: liTop,
                left: liLeft,
                color: liColor
              } = liStyles;
              const paddingLeft = liStyles['padding-left']
              const paddingRight = liStyles['padding-right']
              const paddingTop = liStyles['padding-top']
              const paddingBottom = liStyles['padding-bottom']
              const liAlignItems = liStyles['align-items'];
              const liJustifyContent = liStyles['justify-content'];
              const liFontSize = parseInt(liStyles['font-size'], 10) * .3;
              const liFontWeight = liStyles['font-weight'];
              const liBorderColor = liStyles['border-color'];
              const {x: liX, y: liY, height: h, width: w} = item.getBoundingClientRect()
              // @ts-ignore
              slide.addText(item.textContent.trim(), {
                w: PxToPPT(liWidth),
                h: PxToPPT(liHeight),
                x: PxToPPT(left),
                y: PxToPPT(top) + computeHeight,
                align: justifyContentToPPTalign(liJustifyContent),
                valign: alignItemsToPPTvalign(liAlignItems),
                fontSize: liFontSize,
                color: rgbToHex(liColor),
                bold: liFontWeight === 'bold' || liFontWeight > 400,
                bullet: text.startsWith('<li'),
                margin: [PxToPPT(paddingTop), PxToPPT(paddingRight), PxToPPT(paddingBottom), PxToPPT(paddingLeft)],
              });
              computeHeight += PxToPPT(liHeight);
            })

          } else {
            slide.addText(text, {
              w: PxToPPT(width),
              h: PxToPPT(height),
              x: PxToPPT(left),
              y: PxToPPT(top),
              align: justifyContentToPPTalign(justifyContent),
              valign: alignItemsToPPTvalign(alignItems),
              fontSize,
              color: rgbToHex(color),
              bold: fontWeight === 'bold' || fontWeight > 400,
            });
          }

        } else if (attributes['type'] === 'image') {
          slide.addImage({
            path: attributes['src'],
            x: PxToPPT(left),
            y: PxToPPT(top),
            w: PxToPPT(width),
            h: PxToPPT(height),
          });
        } else if (attributes['type'] === 'table') {
          try {
            const tableData: any = []; // 假设您需要从表格中提取数据
            const rows = child.getElementsByTagName('tr');
            for (let row of rows) {
              const rowData = [];
              const cells = row.getElementsByTagName('td');
              for (let cell of cells) {
                rowData.push(cell.innerText);
              }
              tableData.push(rowData);
            }
            slide.addTable(tableData, {
              w: PxToPPT(width),
              h: PxToPPT(height),
              x: PxToPPT(left),
              y: PxToPPT(top),
              align: justifyContentToPPTalign(justifyContent),
              valign: alignItemsToPPTvalign(alignItems),
              fontSize,
              color: rgbToHex(color),
              bold: fontWeight === 'bold' || fontWeight > 400,
              border: {color: rgbToHex(borderColor)},
            });
          } catch (error) {
            console.error('Error parsing table data:', error);
          }
        }
      }
    });
  })
  return ppt;
}


/** 像素转为 16:9 比例的 pptx 页宽 */
function PxToPPT(pxVal: string): number {
  return parseInt(pxVal, 10) / 192;
}

/** 垂直对齐方式转换为 pptx 的 valign */
function alignItemsToPPTvalign(val: string): any {
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
function justifyContentToPPTalign(val: string): any {
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

/** rgb 转 hex */
function rgbToHex(rgb: string) {
  const result = rgb.match(/\d+/g)?.map(x => {
    const hex = parseInt(x).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  });
  return `#${result?.join('')}`;
}

function traverseElements(element: Element): { styles: any, attributes: any } {
  // 获取元素的属性
  const attributes: any = {};
  Array.from(element.attributes).forEach(attr => attributes[attr.name] = attr.value);
  // 获取元素的样式
  const computedStyle = window.getComputedStyle(element);
  const styles: any = {};
  Array.from(computedStyle).forEach(style => styles[style] = computedStyle.getPropertyValue(style));
  return {styles, attributes}
}
