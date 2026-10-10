/**
 * Page geometry and coordinate transformation between screen canvas (top-left origin)
 * and PDF user space (bottom-left origin).
 */

import { PdfRect } from "./pdf-annotation";

export type PdfRotation = 0 | 90 | 180 | 270;

export interface PdfPageInfo {
  pageIndex: number; // 0-based
  pageNumber: number; // 1-based
  width: number; // PDF points (1/72 inch)
  height: number;
  rotation: PdfRotation;
  scale: number; // Rendering zoom factor
}

/**
 * Coordinate Converter between:
 * - Screen space: CSS pixels within page container, origin at TOP-LEFT.
 * - PDF User space: Points, origin at BOTTOM-LEFT.
 */
export class PdfCoordinateTransformer {
  constructor(
    public readonly pageWidth: number,
    public readonly pageHeight: number,
    public readonly rotation: PdfRotation = 0
  ) {}

  /**
   * Converts a screen coordinate (x, y) relative to page top-left to PDF user space (px, py).
   */
  screenToPdf(screenX: number, screenY: number, scale: number): { x: number; y: number } {
    const unscaledX = screenX / scale;
    const unscaledY = screenY / scale;

    switch (this.rotation) {
      case 90:
        return {
          x: unscaledY,
          y: unscaledX,
        };
      case 180:
        return {
          x: this.pageWidth - unscaledX,
          y: unscaledY,
        };
      case 270:
        return {
          x: this.pageHeight - unscaledY,
          y: this.pageWidth - unscaledX,
        };
      case 0:
      default:
        return {
          x: unscaledX,
          y: this.pageHeight - unscaledY,
        };
    }
  }

  /**
   * Converts a PDF user space coordinate (px, py) to screen space (sx, sy) relative to page top-left.
   */
  pdfToScreen(pdfX: number, pdfY: number, scale: number): { x: number; y: number } {
    let unscaledX: number;
    let unscaledY: number;

    switch (this.rotation) {
      case 90:
        unscaledX = pdfY;
        unscaledY = pdfX;
        break;
      case 180:
        unscaledX = this.pageWidth - pdfX;
        unscaledY = pdfY;
        break;
      case 270:
        unscaledX = this.pageWidth - pdfY;
        unscaledY = this.pageHeight - pdfX;
        break;
      case 0:
      default:
        unscaledX = pdfX;
        unscaledY = this.pageHeight - pdfY;
        break;
    }

    return {
      x: unscaledX * scale,
      y: unscaledY * scale,
    };
  }

  /**
   * Converts a PDF rect [llx, lly, urx, ury] to screen bounds { left, top, width, height }.
   */
  pdfRectToScreenRect(
    rect: PdfRect,
    scale: number
  ): { left: number; top: number; width: number; height: number } {
    const [llx, lly, urx, ury] = rect;
    const p1 = this.pdfToScreen(llx, ury, scale); // top-left in screen
    const p2 = this.pdfToScreen(urx, lly, scale); // bottom-right in screen

    const left = Math.min(p1.x, p2.x);
    const top = Math.min(p1.y, p2.y);
    const width = Math.abs(p2.x - p1.x);
    const height = Math.abs(p2.y - p1.y);

    return { left, top, width, height };
  }

  /**
   * Converts screen bounds { left, top, width, height } to PDF rect [llx, lly, urx, ury].
   */
  screenRectToPdfRect(
    screenRect: { left: number; top: number; width: number; height: number },
    scale: number
  ): PdfRect {
    const pTopLeft = this.screenToPdf(screenRect.left, screenRect.top, scale);
    const pBottomRight = this.screenToPdf(
      screenRect.left + screenRect.width,
      screenRect.top + screenRect.height,
      scale
    );

    const llx = Math.min(pTopLeft.x, pBottomRight.x);
    const lly = Math.min(pTopLeft.y, pBottomRight.y);
    const urx = Math.max(pTopLeft.x, pBottomRight.x);
    const ury = Math.max(pTopLeft.y, pBottomRight.y);

    return [llx, lly, urx, ury];
  }
}
