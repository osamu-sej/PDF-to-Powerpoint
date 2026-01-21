import { ProcessedPage, TextItem, ImageItem } from '../types';

declare global {
  interface Window {
    pdfjsLib: any;
  }
}

// Helper to multiply 2D matrices (3x3 used for 2D affine transforms)
// [a b 0]
// [c d 0]
// [e f 1]
// Stored as [a, b, c, d, e, f]
const multiplyMatrix = (m1: number[], m2: number[]) => {
  const [a1, b1, c1, d1, e1, f1] = m1;
  const [a2, b2, c2, d2, e2, f2] = m2;

  return [
    a1 * a2 + b1 * c2,       // a
    a1 * b2 + b1 * d2,       // b
    c1 * a2 + d1 * c2,       // c
    c1 * b2 + d1 * d2,       // d
    e1 * a2 + f1 * c2 + e2,  // e
    e1 * b2 + f1 * d2 + f2,  // f
  ];
};

const IDENTITY_MATRIX = [1, 0, 0, 1, 0, 0];

export const convertPdfToImages = async (
  file: File,
  onProgress: (progress: number) => void
): Promise<ProcessedPage[]> => {
  const fileArrayBuffer = await file.arrayBuffer();
  
  const loadingTask = window.pdfjsLib.getDocument(fileArrayBuffer);
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;
  const processedPages: ProcessedPage[] = [];

  const BATCH_SIZE = 3; 

  for (let i = 1; i <= numPages; i += BATCH_SIZE) {
    const batchPromises = [];
    
    for (let j = 0; j < BATCH_SIZE && i + j <= numPages; j++) {
      const pageNum = i + j;
      batchPromises.push(processPage(pdf, pageNum));
    }

    const batchResults = await Promise.all(batchPromises);
    processedPages.push(...batchResults);
    
    const progress = Math.min(100, Math.round((processedPages.length / numPages) * 100));
    onProgress(progress);
  }

  return processedPages.sort((a, b) => a.pageNumber - b.pageNumber);
};

const processPage = async (pdf: any, pageNumber: number): Promise<ProcessedPage> => {
  const page = await pdf.getPage(pageNumber);
  
  // 1. Setup Layout Dimensions (Scale 1.0 = 72 DPI standard PDF units)
  const viewportStandard = page.getViewport({ scale: 1.0 });
  const widthInches = viewportStandard.width / 72;
  const heightInches = viewportStandard.height / 72;

  // 2. Extract Text Content (Editable Text)
  const textContent = await page.getTextContent();
  const textItems: TextItem[] = [];

  if (textContent && textContent.items) {
    for (const item of textContent.items) {
      if (!item.str || item.str.trim().length === 0) continue;

      const tx = item.transform;
      const fontSize = Math.sqrt(tx[0] * tx[0] + tx[1] * tx[1]);
      const [xPx, yPx] = viewportStandard.convertToViewportPoint(tx[4], tx[5]);

      const xInch = xPx / 72;
      const yInch = (yPx - (fontSize * 0.8)) / 72;
      const wInch = item.width / 72;
      
      textItems.push({
        text: item.str,
        x: xInch,
        y: yInch,
        w: wInch > 0 ? wInch : 1, 
        h: fontSize / 72 * 1.2,
        fontSize: fontSize,
        fontFace: 'Meiryo UI',
        color: '000000',
        rotation: 0 
      });
    }
  }

  // 3. Extract Raster Images & Vectors
  const operatorList = await page.getOperatorList();
  const imageItems: ImageItem[] = [];
  
  let transformStack: number[][] = [IDENTITY_MATRIX];
  let currentMatrix = IDENTITY_MATRIX;

  const OPS = window.pdfjsLib.OPS;

  const convertToInches = (x: number, y: number) => {
    const [vx, vy] = viewportStandard.convertToViewportPoint(x, y);
    return { x: vx / 72, y: vy / 72 };
  };

  const fnArray = operatorList.fn || [];
  const argsArray = operatorList.args || [];

  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i];
    const args = argsArray[i];

    if (fn === OPS.save) {
      transformStack.push(currentMatrix);
    } else if (fn === OPS.restore) {
      if (transformStack.length > 0) {
        currentMatrix = transformStack.pop()!;
      }
    } else if (fn === OPS.transform) {
      currentMatrix = multiplyMatrix(currentMatrix, args);
    } else if (fn === OPS.paintImageXObject || fn === OPS.paintInlineImageXObject) {
      const imgName = args[0];
      
      try {
        let imgObj = null;
        if (page.objs) imgObj = page.objs.get(imgName);
        if (!imgObj && page.commonObjs) imgObj = page.commonObjs.get(imgName);
        
        if (imgObj && !imgObj.data && typeof imgObj.then === 'function') {
             continue;
        }

        if (imgObj && imgObj.width && imgObj.height) {
           const applyMat = (m: number[], p: [number, number]) => {
             return [
               m[0]*p[0] + m[2]*p[1] + m[4],
               m[1]*p[0] + m[3]*p[1] + m[5]
             ];
           };

           const p00 = applyMat(currentMatrix, [0, 0]);
           const p10 = applyMat(currentMatrix, [1, 0]);
           const p01 = applyMat(currentMatrix, [0, 1]);
           const p11 = applyMat(currentMatrix, [1, 1]);
           
           const c00 = convertToInches(p00[0], p00[1]);
           const c10 = convertToInches(p10[0], p10[1]);
           const c01 = convertToInches(p01[0], p01[1]);
           const c11 = convertToInches(p11[0], p11[1]);

           const xs = [c00.x, c10.x, c01.x, c11.x];
           const ys = [c00.y, c10.y, c01.y, c11.y];
           
           const minX = Math.min(...xs);
           const minY = Math.min(...ys);
           const maxX = Math.max(...xs);
           const maxY = Math.max(...ys);
           
           const w = maxX - minX;
           const h = maxY - minY;

           const tmpCanvas = document.createElement('canvas');
           tmpCanvas.width = imgObj.width;
           tmpCanvas.height = imgObj.height;
           const tmpCtx = tmpCanvas.getContext('2d', { willReadFrequently: true });
           
           if (tmpCtx) {
             let validData = false;
             const arr = new Uint8ClampedArray(imgObj.width * imgObj.height * 4);
             
             if (imgObj.kind === window.pdfjsLib.ImageKind.RGB_24BPP && imgObj.data && imgObj.data.length) {
                 const data = imgObj.data;
                 let j = 0;
                 for (let k = 0; k < data.length; k += 3) {
                     arr[j++] = data[k];
                     arr[j++] = data[k + 1];
                     arr[j++] = data[k + 2];
                     arr[j++] = 255;
                 }
                 validData = true;
             } else if (imgObj.kind === window.pdfjsLib.ImageKind.RGBA_32BPP && imgObj.data) {
                  arr.set(imgObj.data);
                  validData = true;
             } else if (imgObj.bitmap) {
                 validData = true;
             }
             
             if (validData) {
                 if (imgObj.bitmap) {
                     tmpCtx.drawImage(imgObj.bitmap, 0, 0);
                 } else {
                     const id = new ImageData(arr, imgObj.width, imgObj.height);
                     tmpCtx.putImageData(id, 0, 0);
                 }

                 imageItems.push({
                   data: tmpCanvas.toDataURL('image/png'),
                   x: minX,
                   y: minY,
                   w: w,
                   h: h
                 });
             }
           }
        }
      } catch (e) {
        console.warn("Failed to extract image", e);
      }
    }
  }

  // 4. Render Background Image
  // We allow drawImage here to ensure all visual elements are captured in the background.
  // This might duplicate images that were successfully extracted above, but it prevents data loss.
  const scale = 2.0; 
  const viewportHighRes = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d', { willReadFrequently: true });
  
  if (!context) throw new Error('Could not get canvas context');

  canvas.height = viewportHighRes.height;
  canvas.width = viewportHighRes.width;

  const proxyContext = new Proxy(context, {
    get(target, prop: string | symbol) {
      // Block TEXT only. We want text to be editable (extracted in step 2), not baked in.
      if (prop === 'fillText' || prop === 'strokeText') {
        return () => {};
      }
      
      const value = (target as any)[prop];
      if (typeof value === 'function') {
        return value.bind(target);
      }
      return value;
    },
    set(target, prop: string | symbol, value: any) {
      (target as any)[prop] = value;
      return true;
    }
  });

  await page.render({
    canvasContext: proxyContext as any,
    viewport: viewportHighRes,
  }).promise;

  const imageData = canvas.toDataURL('image/jpeg', 0.85);

  return {
    pageNumber,
    imageData, 
    width: widthInches,
    height: heightInches,
    textItems: textItems || [],
    images: imageItems || [],
  };
};