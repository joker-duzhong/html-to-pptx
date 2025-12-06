import { html2pptx } from "./analyze";

export type WRITE_OUTPUT_TYPE = "arraybuffer" | "base64" | "binarystring" | "blob" | "nodebuffer" | "uint8array" | "STREAM";

export function exportHtmlToPpt(pageClassName: string = "page", outputType: WRITE_OUTPUT_TYPE = "blob"): Promise<string | ArrayBuffer | Blob | Uint8Array> {
  return new Promise(async (resolve, reject) => {
    try {
      const PptxGenJSIns = await html2pptx(pageClassName);
      const result = PptxGenJSIns.write({ outputType });
      resolve(result);
    } catch (error) {
      console.error("PPT Generation Error:", error);
      reject(error);
    }
  });
}

export function downloadHtmlToPpt(pageClassName: string = "page", fileName: string = "presentation"): Promise<void> {
  return new Promise(async (resolve, reject) => {
    try {
      // 使用 writeFile 自动触发浏览器下载
      const PptxGenJSIns = await html2pptx(pageClassName);
      PptxGenJSIns.writeFile({ fileName: fileName + ".pptx" });
      resolve();
    } catch (error) {
      console.error("PPT Generation Error:", error);
      reject(error);
    }
  });
}
