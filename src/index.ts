import { html2pptx } from "./analyze";

export type WRITE_OUTPUT_TYPE = "arraybuffer" | "base64" | "binarystring" | "blob" | "nodebuffer" | "uint8array" | "STREAM";

export function exportHtmlToPpt(pageClassName: string = "page", outputType: WRITE_OUTPUT_TYPE = "blob"): Promise<string | ArrayBuffer | Blob | Uint8Array> {
  return html2pptx(pageClassName).write({ outputType });
}

export function downloadHtmlToPpt(pageClassName: string = "page", fileName: string = "hope"): Promise<void> {
  return new Promise(async (resolve, reject) => {
    try {
      await html2pptx(pageClassName).writeFile({ fileName: fileName + ".pptx" });
      resolve();
    } catch (error) {
      console.error("PPT Generation Error:", error);
      reject(error);
    }
  });
}