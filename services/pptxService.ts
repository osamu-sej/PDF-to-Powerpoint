import { ProcessedPage } from '../types';

declare global {
  interface Window {
    PptxGenJS: any;
  }
}

export const generatePptxFile = async (
  pages: ProcessedPage[],
  fileName: string
): Promise<void> => {
  const pptx = new window.PptxGenJS();

  // Set Metadata
  pptx.title = fileName.replace('.pdf', '');
  pptx.subject = 'Converted from PDF via AI-Powered Converter';

  // FIX: Set Layout to match the first page's dimensions.
  // This prevents the "143%" zoom effect caused by fitting custom PDF sizes (like A4) 
  // into default PowerPoint slides.
  if (pages.length > 0) {
    const firstPage = pages[0];
    const layoutName = 'PDF_CUSTOM_LAYOUT';
    
    pptx.defineLayout({
      name: layoutName,
      width: firstPage.width,
      height: firstPage.height,
    });
    
    pptx.layout = layoutName;
  }

  for (const page of pages) {
    const slide = pptx.addSlide();

    // 1. Add the "Background" Layer (Vectors, Shapes, Colors)
    // This provides the visual base.
    slide.addImage({
      data: page.imageData,
      x: 0,
      y: 0,
      w: page.width,
      h: page.height,
    });
    
    // 2. Add Extracted Raster Images (Movable Figures)
    // Placed on top of the background for editability.
    for (const img of page.images) {
      slide.addImage({
        data: img.data,
        x: img.x,
        y: img.y,
        w: img.w,
        h: img.h,
      });
    }

    // 3. Add Editable Text Boxes on top
    for (const item of page.textItems) {
      slide.addText(item.text, {
        x: item.x,
        y: item.y,
        w: item.w + 0.1, // Add slight buffer to prevent wrapping
        h: item.h,
        fontSize: item.fontSize,
        fontFace: item.fontFace,
        color: item.color,
        // Align properties to try to match PDF look
        align: 'left',
        valign: 'top',
        margin: 0, // Minimize internal margin to match PDF positioning
      });
    }

    // 4. Add AI generated notes if available
    if (page.aiNotes) {
      slide.addNotes(page.aiNotes);
    }
  }

  await pptx.writeFile({ fileName: `${fileName.replace('.pdf', '')}.pptx` });
};