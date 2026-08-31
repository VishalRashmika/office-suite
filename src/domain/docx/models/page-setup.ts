import { NS } from "../../../core/utils/xml";

export interface DocxPageSetup {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  contentWidth: number;
  contentHeight: number;
  orientation: "portrait" | "landscape";
}

/**
 * Calculates page dimensions and margins from a WordprocessingML <w:sectPr> element.
 */
export function parseDocxPageSetup(sectPr: Element | null): DocxPageSetup {
  // Default standard Letter size (8.5 x 11 in, 1440 dxa / 1in margins = 96px)
  let pageWidth = 816;
  let pageHeight = 1056;
  let marginTop = 72;
  let marginBottom = 72;
  let marginLeft = 72;
  let marginRight = 72;
  let orientation: "portrait" | "landscape" = "portrait";

  if (sectPr) {
    const pgSz = sectPr.getElementsByTagNameNS(NS.w, "pgSz")[0] || sectPr.querySelector("pgSz");
    if (pgSz) {
      const wDxa = parseInt(pgSz.getAttributeNS(NS.w, "w") || pgSz.getAttribute("w:w") || "12240", 10);
      const hDxa = parseInt(pgSz.getAttributeNS(NS.w, "h") || pgSz.getAttribute("w:h") || "15840", 10);
      const orient = pgSz.getAttributeNS(NS.w, "orient") || pgSz.getAttribute("w:orient");
      if (wDxa > 0) pageWidth = Math.round(wDxa / 15);
      if (hDxa > 0) pageHeight = Math.round(hDxa / 15);
      if (orient === "landscape") orientation = "landscape";
    }

    const pgMar = sectPr.getElementsByTagNameNS(NS.w, "pgMar")[0] || sectPr.querySelector("pgMar");
    if (pgMar) {
      const topDxa = parseInt(pgMar.getAttributeNS(NS.w, "top") || pgMar.getAttribute("w:top") || "1080", 10);
      const bottomDxa = parseInt(pgMar.getAttributeNS(NS.w, "bottom") || pgMar.getAttribute("w:bottom") || "1080", 10);
      const leftDxa = parseInt(pgMar.getAttributeNS(NS.w, "left") || pgMar.getAttribute("w:left") || "1080", 10);
      const rightDxa = parseInt(pgMar.getAttributeNS(NS.w, "right") || pgMar.getAttribute("w:right") || "1080", 10);
      if (topDxa > 0) marginTop = Math.max(36, Math.min(144, Math.round(topDxa / 15)));
      if (bottomDxa > 0) marginBottom = Math.max(36, Math.min(144, Math.round(bottomDxa / 15)));
      if (leftDxa > 0) marginLeft = Math.max(36, Math.min(144, Math.round(leftDxa / 15)));
      if (rightDxa > 0) marginRight = Math.max(36, Math.min(144, Math.round(rightDxa / 15)));
    }
  }

  const contentWidth = Math.max(200, pageWidth - marginLeft - marginRight);
  const contentHeight = Math.max(300, pageHeight - marginTop - marginBottom);

  return {
    pageWidth,
    pageHeight,
    marginTop,
    marginBottom,
    marginLeft,
    marginRight,
    contentWidth,
    contentHeight,
    orientation,
  };
}
