/**
 * Domain services for sorting, grouping, filtering, and color-organizing PDF annotations.
 * Clones Zotero's left sidebar annotations organizing capabilities:
 * - Sort by Color (grouped by semantic tags like Green = "New Info", Red = "Important Points")
 * - Sort by Document Page / Position (reading order)
 * - Sort by Date (newest / oldest)
 * - Sort by Annotation Type
 * - Quick filter by Color palette, Semantic Tag, and Type
 * - Full-text search with page number matching
 */

import {
  PdfAnnotation,
  PdfAnnotationType,
  PdfRgbColor,
  PDF_COLOR_PRESETS,
  hexToRgb,
  rgbToHex,
} from "../models/pdf-annotation";
import {
  PdfColorSchemeProfile,
  findTagForColor,
  normalizeHexColor,
} from "../models/pdf-color-scheme";

export type AnnotSortField = "page" | "color" | "date" | "type";
export type AnnotSortDirection = "asc" | "desc";

export interface AnnotSortOptions {
  field: AnnotSortField;
  direction?: AnnotSortDirection;
}

export interface AnnotFilterOptions {
  searchQuery?: string;
  type?: "all" | PdfAnnotationType | "shape";
  colorHex?: string; // specific normalized hex or 'all'
  tag?: string; // specific tag name or 'all'
}

export interface PdfColorGroup {
  key: string;
  hex: string;
  name: string;
  tag?: string;
  rgb: PdfRgbColor;
  annotations: PdfAnnotation[];
}

export interface PdfTypeGroup {
  type: string;
  label: string;
  icon: string;
  annotations: PdfAnnotation[];
}

export interface AnnotationColorInfo {
  hex: string;
  name: string;
  tag?: string;
  rgb: PdfRgbColor;
}

/**
 * Resolves full color metadata (hex, name, semantic tag, rgb) for any annotation
 * using the document's active color scheme profile and standard Zotero presets.
 */
export function getAnnotationColorInfo(
  annotation: PdfAnnotation,
  scheme?: PdfColorSchemeProfile
): AnnotationColorInfo {
  const tagged = findTagForColor(scheme, annotation.color);
  if (tagged) {
    return {
      hex: normalizeHexColor(tagged.hex),
      name: tagged.name,
      tag: tagged.tag || undefined,
      rgb: tagged.rgb || hexToRgb(tagged.hex),
    };
  }

  const hex = normalizeHexColor(rgbToHex(annotation.color));
  const preset = PDF_COLOR_PRESETS.find(
    (p) => normalizeHexColor(p.hex) === hex
  );
  if (preset) {
    return {
      hex: preset.hex,
      name: preset.name,
      rgb: preset.rgb,
    };
  }

  return {
    hex,
    name: "Custom Color",
    rgb: annotation.color || hexToRgb(hex),
  };
}

/**
 * Filters annotations by search query, annotation type, color, and semantic tag.
 */
export function filterAnnotations(
  annotations: PdfAnnotation[],
  options: AnnotFilterOptions,
  scheme?: PdfColorSchemeProfile
): PdfAnnotation[] {
  let result = [...annotations];

  // 1. Type Filter
  if (options.type && options.type !== "all") {
    if (options.type === "shape") {
      result = result.filter(
        (a) => a.type === "square" || a.type === "circle" || a.type === "line"
      );
    } else {
      result = result.filter((a) => a.type === options.type);
    }
  }

  // 2. Color Filter (by Hex)
  if (options.colorHex && options.colorHex !== "all") {
    const targetHex = normalizeHexColor(options.colorHex);
    result = result.filter((a) => {
      const info = getAnnotationColorInfo(a, scheme);
      return normalizeHexColor(info.hex) === targetHex;
    });
  }

  // 3. Semantic Tag Filter
  if (options.tag && options.tag !== "all") {
    result = result.filter((a) => {
      const info = getAnnotationColorInfo(a, scheme);
      return info.tag === options.tag;
    });
  }

  // 4. Search Query Filter
  if (options.searchQuery && options.searchQuery.trim().length > 0) {
    const q = options.searchQuery.toLowerCase().trim();

    // Check if query is searching for page (e.g. "p. 3", "page 3", or "3")
    const pageMatch = q.match(/^(?:p\.?|page)?\s*(\d+)$/i);
    const targetPageNum = pageMatch ? parseInt(pageMatch[1], 10) : null;

    result = result.filter((a) => {
      const hlText = ((a as any).highlightedText || "").toLowerCase();
      const contents = (a.contents || "").toLowerCase();
      const author = (a.author || "").toLowerCase();
      const info = getAnnotationColorInfo(a, scheme);
      const tagName = (info.tag || "").toLowerCase();
      const colorName = info.name.toLowerCase();
      const linkedNote = (a.linkedNote || "").toLowerCase();
      const typeLabel = formatTypeLabel(a.type).toLowerCase();

      // Check highlighted text, note/comment, author, tag, color name, linked note, or type
      if (
        hlText.includes(q) ||
        contents.includes(q) ||
        author.includes(q) ||
        tagName.includes(q) ||
        colorName.includes(q) ||
        linkedNote.includes(q) ||
        typeLabel.includes(q)
      ) {
        return true;
      }

      // Check page number match
      if (targetPageNum !== null && a.pageIndex + 1 === targetPageNum) {
        return true;
      }

      return false;
    });
  }

  return result;
}

/**
 * Sorts annotations according to specified sort options (page, color, date, or type).
 */
export function sortAnnotations(
  annotations: PdfAnnotation[],
  options: AnnotSortOptions,
  scheme?: PdfColorSchemeProfile
): PdfAnnotation[] {
  const { field, direction = "asc" } = options;
  const list = [...annotations];

  // Secondary tie-breaker: page and reading position (top-to-bottom, left-to-right)
  const compareByPosition = (a: PdfAnnotation, b: PdfAnnotation): number => {
    if (a.pageIndex !== b.pageIndex) {
      return a.pageIndex - b.pageIndex;
    }
    // In PDF coordinate space, Y=0 is bottom, so highest Y is top of page
    const yA = a.rect ? a.rect[3] : 0;
    const yB = b.rect ? b.rect[3] : 0;
    if (Math.abs(yB - yA) > 1) {
      return yB - yA;
    }
    const xA = a.rect ? a.rect[0] : 0;
    const xB = b.rect ? b.rect[0] : 0;
    return xA - xB;
  };

  switch (field) {
    case "page": {
      list.sort((a, b) => {
        const cmp = compareByPosition(a, b);
        return direction === "asc" ? cmp : -cmp;
      });
      break;
    }

    case "date": {
      list.sort((a, b) => {
        const timeA = a.modifiedAt
          ? a.modifiedAt.getTime()
          : a.createdAt
          ? a.createdAt.getTime()
          : 0;
        const timeB = b.modifiedAt
          ? b.modifiedAt.getTime()
          : b.createdAt
          ? b.createdAt.getTime()
          : 0;

        if (timeA !== timeB) {
          // Default date sort: descending (newest first)
          return direction === "asc" ? timeA - timeB : timeB - timeA;
        }
        return compareByPosition(a, b);
      });
      break;
    }

    case "color": {
      // Build color order from active scheme
      const colorOrderMap = new Map<string, number>();
      if (scheme && scheme.colors) {
        scheme.colors.forEach((c, idx) => {
          colorOrderMap.set(normalizeHexColor(c.hex), idx);
        });
      }
      PDF_COLOR_PRESETS.forEach((p, idx) => {
        const hex = normalizeHexColor(p.hex);
        if (!colorOrderMap.has(hex)) {
          colorOrderMap.set(hex, 100 + idx);
        }
      });

      list.sort((a, b) => {
        const infoA = getAnnotationColorInfo(a, scheme);
        const infoB = getAnnotationColorInfo(b, scheme);
        const rankA = colorOrderMap.get(infoA.hex) ?? 999;
        const rankB = colorOrderMap.get(infoB.hex) ?? 999;

        if (rankA !== rankB) {
          return direction === "asc" ? rankA - rankB : rankB - rankA;
        }
        return compareByPosition(a, b);
      });
      break;
    }

    case "type": {
      const typeOrder: Record<string, number> = {
        highlight: 1,
        underline: 2,
        text: 3,
        freetext: 4,
        ink: 5,
        square: 6,
        circle: 7,
        line: 8,
      };

      list.sort((a, b) => {
        const orderA = typeOrder[a.type] ?? 99;
        const orderB = typeOrder[b.type] ?? 99;
        if (orderA !== orderB) {
          return direction === "asc" ? orderA - orderB : orderB - orderA;
        }
        return compareByPosition(a, b);
      });
      break;
    }
  }

  return list;
}

/**
 * Groups annotations by color, resolving semantic tags and color names.
 * Ideal for Zotero's color-categorized display mode.
 */
export function groupAnnotationsByColor(
  annotations: PdfAnnotation[],
  scheme?: PdfColorSchemeProfile,
  direction: AnnotSortDirection = "asc"
): PdfColorGroup[] {
  const groupsMap = new Map<string, PdfColorGroup>();

  for (const annot of annotations) {
    const info = getAnnotationColorInfo(annot, scheme);
    const key = info.hex;

    let group = groupsMap.get(key);
    if (!group) {
      group = {
        key,
        hex: info.hex,
        name: info.name,
        tag: info.tag,
        rgb: info.rgb,
        annotations: [],
      };
      groupsMap.set(key, group);
    }
    group.annotations.push(annot);
  }

  // Sort annotations within each group by page position
  for (const group of groupsMap.values()) {
    group.annotations = sortAnnotations(group.annotations, {
      field: "page",
      direction: "asc",
    });
  }

  // Determine ordering of color groups according to active scheme
  const colorOrderMap = new Map<string, number>();
  if (scheme && scheme.colors) {
    scheme.colors.forEach((c, idx) => {
      colorOrderMap.set(normalizeHexColor(c.hex), idx);
    });
  }
  PDF_COLOR_PRESETS.forEach((p, idx) => {
    const hex = normalizeHexColor(p.hex);
    if (!colorOrderMap.has(hex)) {
      colorOrderMap.set(hex, 100 + idx);
    }
  });

  const groups = Array.from(groupsMap.values());
  groups.sort((a, b) => {
    const rankA = colorOrderMap.get(a.hex) ?? 999;
    const rankB = colorOrderMap.get(b.hex) ?? 999;
    return direction === "asc" ? rankA - rankB : rankB - rankA;
  });

  return groups;
}

/**
 * Groups annotations by type (Highlights, Underlines, Sticky Notes, FreeText, Drawings, Shapes).
 */
export function groupAnnotationsByType(
  annotations: PdfAnnotation[],
  direction: AnnotSortDirection = "asc"
): PdfTypeGroup[] {
  const defs: { type: string; label: string; icon: string; match: (a: PdfAnnotation) => boolean }[] = [
    { type: "highlight", label: "Highlights", icon: "highlighter", match: (a) => a.type === "highlight" },
    { type: "underline", label: "Underlines", icon: "underline", match: (a) => a.type === "underline" },
    { type: "text", label: "Sticky Notes", icon: "message-square", match: (a) => a.type === "text" },
    { type: "freetext", label: "Text Boxes", icon: "type", match: (a) => a.type === "freetext" },
    { type: "ink", label: "Drawings", icon: "pen-tool", match: (a) => a.type === "ink" },
    {
      type: "shape",
      label: "Shapes & Lines",
      icon: "square",
      match: (a) => a.type === "square" || a.type === "circle" || a.type === "line",
    },
  ];

  const result: PdfTypeGroup[] = [];

  for (const def of defs) {
    const matching = annotations.filter(def.match);
    if (matching.length > 0) {
      result.push({
        type: def.type,
        label: def.label,
        icon: def.icon,
        annotations: sortAnnotations(matching, { field: "page", direction: "asc" }),
      });
    }
  }

  if (direction === "desc") {
    result.reverse();
  }

  return result;
}

/**
 * Formats annotation type for user display.
 */
export function formatTypeLabel(type: PdfAnnotationType | string): string {
  switch (type) {
    case "highlight":
      return "Highlight";
    case "underline":
      return "Underline";
    case "text":
      return "Sticky Note";
    case "ink":
      return "Drawing";
    case "freetext":
      return "Text Box";
    case "square":
      return "Rectangle";
    case "circle":
      return "Circle";
    case "line":
      return "Line";
    default:
      return "Annotation";
  }
}

/**
 * Human-friendly date formatting for annotations.
 */
export function formatAnnotationDate(date?: Date | number | string): string {
  if (!date) return "";
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return "";

  const now = new Date();
  const diffMs = now.getTime() - d.getTime();
  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHours = Math.floor(diffMin / 60);

  if (diffSec < 60 && diffSec >= 0) {
    return "Just now";
  }
  if (diffMin < 60 && diffMin > 0) {
    return `${diffMin}m ago`;
  }

  const isToday =
    d.getDate() === now.getDate() &&
    d.getMonth() === now.getMonth() &&
    d.getFullYear() === now.getFullYear();

  const timeStr = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

  if (isToday) {
    return `Today, ${timeStr}`;
  }

  const yesterday = new Date(now);
  yesterday.setDate(now.getDate() - 1);
  const isYesterday =
    d.getDate() === yesterday.getDate() &&
    d.getMonth() === yesterday.getMonth() &&
    d.getFullYear() === yesterday.getFullYear();

  if (isYesterday) {
    return `Yesterday, ${timeStr}`;
  }

  return d.toLocaleDateString([], {
    month: "short",
    day: "numeric",
    year: d.getFullYear() !== now.getFullYear() ? "numeric" : undefined,
  });
}
