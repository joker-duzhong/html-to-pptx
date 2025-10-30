import { html2pptx } from './analyze'

/** 导出类型 */
export type WRITE_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array' | 'STREAM';

/**
 * [说明文件](../README.md)
 * 导出页面dom为PPT (支持 text | image | table)
 * 请在页面中的 dom元素添加以下属性 type: text|image|table
 * @param pageClass 每一页的页面类名  (默认为 page)
 * @param outputType 输出文件类型 
 * @returns ppt文件
 */
export function exportHtmlToPpt(pageClassName: string = 'page', outputType: WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Uint8Array> {
  return html2pptx(pageClassName)?.write({ outputType });
}

/**
 * [说明文件](../README.md)
 * 下载页面dom转为为PPT (支持 text | image | table)
 * 请在页面中的 dom元素添加以下属性 type: text|image|table
 * @param pageClass 每一页的页面类名  (默认为 page)
 * @param fileName 导出的ppt文件名
 * @returns 
 */
export function downloadHtmlToPpt(pageClassName: string = 'page', fileName: string = 'hope'): Promise<string | ArrayBuffer | Blob | Uint8Array> {
  return html2pptx(pageClassName).writeFile({ fileName: fileName + '.pptx' });
}