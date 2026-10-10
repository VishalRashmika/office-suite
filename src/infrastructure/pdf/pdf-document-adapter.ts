/**
 * PDF Document Adapter wrapping pdf-lib for lossless document editing,
 * annotation synchronization, and page manipulation.
 */

import { PDFDocument, PDFPage, degrees, PDFObjectStreamParser, PDFRef } from "pdf-lib";
import {
  PdfAnnotation,
  PdfPageInfo,
  PdfRotation,
  PdfAnnotationParser,
  PdfAnnotationSerializer,
} from "../../domain/pdf";

// Resilient Object Stream patch for pdf-lib:
// Prevents an entire object stream (and all its annotations) from being discarded
// if an individual object encounters a syntax warning (e.g. literal string unbalanced parens).
let isObjectStreamParserPatched = false;
function patchPdfLibObjectStreamParser(): void {
  if (isObjectStreamParserPatched) return;
  isObjectStreamParserPatched = true;

  try {
    const proto = (PDFObjectStreamParser as any)?.prototype;
    if (proto && typeof proto.parseIntoContext === "function") {
      proto.parseIntoContext = async function () {
        if (this.alreadyParsed) return;
        this.alreadyParsed = true;
        const offsetsAndObjectNumbers = this.parseOffsetsAndObjectNumbers();
        for (let idx = 0; idx < offsetsAndObjectNumbers.length; idx++) {
          const { objectNumber, offset } = offsetsAndObjectNumbers[idx];
          try {
            this.bytes.moveTo(this.firstOffset + offset);
            const object = this.parseObject();
            const ref = PDFRef.of(objectNumber, 0);
            this.context.assign(ref, object);
          } catch (err) {
            // Gracefully continue to parse remaining objects in the stream
          }
        }
      };
    }
  } catch (err) {
    console.warn("Could not patch PDFObjectStreamParser:", err);
  }
}

export class PdfDocumentAdapter {
  private pdfDoc: PDFDocument;
  private annotations: Map<string, PdfAnnotation> = new Map();
  private originalBytes: ArrayBuffer;
  private isDirty = false;

  private constructor(pdfDoc: PDFDocument, originalBytes: ArrayBuffer) {
    this.pdfDoc = pdfDoc;
    this.originalBytes = originalBytes;
  }

  /**
   * Loads a PDF document from binary bytes.
   */
  static async load(bytes: ArrayBuffer): Promise<PdfDocumentAdapter> {
    patchPdfLibObjectStreamParser();
    const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const adapter = new PdfDocumentAdapter(pdfDoc, bytes);
    adapter.loadExistingAnnotations();
    return adapter;
  }

  /**
   * Scans all pages and parses existing annotations into memory.
   */
  private loadExistingAnnotations(): void {
    this.annotations.clear();
    const pageCount = this.pdfDoc.getPageCount();
    for (let i = 0; i < pageCount; i++) {
      const page = this.pdfDoc.getPage(i);
      const annots = PdfAnnotationParser.parsePageAnnotations(this.pdfDoc, page, i);
      for (const annot of annots) {
        this.annotations.set(annot.id, annot);
      }
    }
  }

  get dirty(): boolean {
    return this.isDirty;
  }

  markClean(): void {
    this.isDirty = false;
  }

  getPageCount(): number {
    return this.pdfDoc.getPageCount();
  }

  getTitle(): string | undefined {
    return this.pdfDoc.getTitle() || undefined;
  }

  getAuthor(): string | undefined {
    return this.pdfDoc.getAuthor() || undefined;
  }

  /**
   * Retrieves geometry and rotation info for a specific page.
   */
  getPageInfo(pageIndex: number, scale = 1.0): PdfPageInfo {
    const page = this.pdfDoc.getPage(pageIndex);
    const { width, height } = page.getSize();
    const rotAngle = (page.getRotation().angle % 360) as PdfRotation;

    return {
      pageIndex,
      pageNumber: pageIndex + 1,
      width,
      height,
      rotation: rotAngle,
      scale,
    };
  }

  /**
   * Retrieves all annotations in the document.
   */
  getAllAnnotations(): PdfAnnotation[] {
    return Array.from(this.annotations.values());
  }

  /**
   * Retrieves annotations belonging to a specific page.
   */
  getPageAnnotations(pageIndex: number): PdfAnnotation[] {
    return Array.from(this.annotations.values()).filter((a) => a.pageIndex === pageIndex);
  }

  /**
   * Retrieves a single annotation by ID.
   */
  getAnnotation(id: string): PdfAnnotation | undefined {
    return this.annotations.get(id);
  }

  /**
   * Registers an existing annotation (e.g. from PDF.js discovery) without marking dirty.
   */
  registerExistingAnnotation(annotation: PdfAnnotation): void {
    if (!this.annotations.has(annotation.id)) {
      this.annotations.set(annotation.id, annotation);
    }
  }

  /**
   * Adds or updates an annotation.
   */
  upsertAnnotation(annotation: PdfAnnotation): void {
    this.annotations.set(annotation.id, {
      ...annotation,
      modifiedAt: new Date(),
    });
    this.isDirty = true;
  }

  /**
   * Deletes an annotation by ID.
   */
  deleteAnnotation(id: string): boolean {
    const existed = this.annotations.delete(id);
    if (existed) {
      this.isDirty = true;
    }
    return existed;
  }

  // ==================== Native PDF Editor Page Operations ====================

  /**
   * Rotates a page by +90, -90, or 180 degrees.
   */
  rotatePage(pageIndex: number, deltaAngle: 90 | -90 | 180): PdfRotation {
    const page = this.pdfDoc.getPage(pageIndex);
    const currentAngle = page.getRotation().angle;
    let newAngle = (currentAngle + deltaAngle) % 360;
    if (newAngle < 0) newAngle += 360;
    page.setRotation(degrees(newAngle));
    this.isDirty = true;
    return newAngle as PdfRotation;
  }

  /**
   * Deletes a page from the document.
   */
  deletePage(pageIndex: number): void {
    if (this.pdfDoc.getPageCount() <= 1) {
      throw new Error("Cannot delete the only page of a PDF document.");
    }

    this.pdfDoc.removePage(pageIndex);

    // Remove annotations on this page and re-index annotations on subsequent pages
    const updated = new Map<string, PdfAnnotation>();
    for (const [id, annot] of this.annotations.entries()) {
      if (annot.pageIndex === pageIndex) {
        continue;
      }
      if (annot.pageIndex > pageIndex) {
        updated.set(id, { ...annot, pageIndex: annot.pageIndex - 1 });
      } else {
        updated.set(id, annot);
      }
    }
    this.annotations = updated;
    this.isDirty = true;
  }

  /**
   * Duplicates an existing page.
   */
  async duplicatePage(pageIndex: number): Promise<void> {
    const [copiedPage] = await this.pdfDoc.copyPages(this.pdfDoc, [pageIndex]);
    const insertIndex = pageIndex + 1;
    this.pdfDoc.insertPage(insertIndex, copiedPage);

    // Shift annotations on subsequent pages
    const updated = new Map<string, PdfAnnotation>();
    for (const [id, annot] of this.annotations.entries()) {
      if (annot.pageIndex >= insertIndex) {
        updated.set(id, { ...annot, pageIndex: annot.pageIndex + 1 });
      } else {
        updated.set(id, annot);
      }
    }
    this.annotations = updated;
    this.isDirty = true;
  }

  /**
   * Inserts a blank page.
   */
  insertBlankPage(pageIndex: number, width = 612, height = 792): void {
    this.pdfDoc.insertPage(pageIndex, [width, height]);

    const updated = new Map<string, PdfAnnotation>();
    for (const [id, annot] of this.annotations.entries()) {
      if (annot.pageIndex >= pageIndex) {
        updated.set(id, { ...annot, pageIndex: annot.pageIndex + 1 });
      } else {
        updated.set(id, annot);
      }
    }
    this.annotations = updated;
    this.isDirty = true;
  }

  // ==================== Serialization & Saving ====================

  /**
   * Serializes in-memory annotations into PDF /Annots dictionaries and saves to Uint8Array.
   */
  async save(): Promise<Uint8Array> {
    const pageCount = this.pdfDoc.getPageCount();

    for (let i = 0; i < pageCount; i++) {
      const page = this.pdfDoc.getPage(i);
      const pageAnnots = this.getPageAnnotations(i);
      PdfAnnotationSerializer.syncPageAnnotations(this.pdfDoc, page, pageAnnots);
    }

    const savedBytes = await this.pdfDoc.save();
    this.originalBytes = savedBytes.buffer.slice(
      savedBytes.byteOffset,
      savedBytes.byteOffset + savedBytes.byteLength
    );
    this.isDirty = false;
    return savedBytes;
  }

  /**
   * Returns current binary buffer representation (e.g. for reloading in PDF.js).
   */
  async getBytes(): Promise<ArrayBuffer> {
    if (this.isDirty) {
      const saved = await this.save();
      return saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer;
    }
    return this.originalBytes;
  }
}
