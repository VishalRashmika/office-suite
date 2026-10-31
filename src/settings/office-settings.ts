/**
 * Plugin configuration and persistent settings for Obsidian Office Suite.
 * Stores global color scheme profiles, default scheme settings, and per-PDF scheme overrides.
 */

import {
  PdfColorSchemeProfile,
  BUILT_IN_PROFILES,
  ZOTERO_DEFAULT_PROFILE,
  ensureValidProfile,
  cloneProfile,
} from "../domain/pdf";

export interface OfficeSuiteSettings {
  pdfDefaultProfileId: string;
  pdfProfiles: PdfColorSchemeProfile[];
  pdfFileSchemes: Record<string, PdfColorSchemeProfile>; // Keyed by file path (e.g. "Papers/AI.pdf")
}

export const DEFAULT_OFFICE_SETTINGS: OfficeSuiteSettings = {
  pdfDefaultProfileId: "zotero-default",
  pdfProfiles: [...BUILT_IN_PROFILES],
  pdfFileSchemes: {},
};

/**
 * Validates and merges loaded data with default settings,
 * ensuring all built-in profiles are present and schemas are non-corrupt.
 */
export function normalizeSettings(loadedData: any): OfficeSuiteSettings {
  const settings: OfficeSuiteSettings = {
    pdfDefaultProfileId: DEFAULT_OFFICE_SETTINGS.pdfDefaultProfileId,
    pdfProfiles: [],
    pdfFileSchemes: {},
  };

  if (!loadedData || typeof loadedData !== "object") {
    settings.pdfProfiles = BUILT_IN_PROFILES.map((p) => cloneProfile(p, p.id, p.name));
    return settings;
  }

  // 1. Default Profile ID
  if (typeof loadedData.pdfDefaultProfileId === "string" && loadedData.pdfDefaultProfileId.trim()) {
    settings.pdfDefaultProfileId = loadedData.pdfDefaultProfileId.trim();
  }

  // 2. Color Profiles
  const profileMap = new Map<string, PdfColorSchemeProfile>();

  // Ensure built-ins exist
  for (const builtIn of BUILT_IN_PROFILES) {
    profileMap.set(builtIn.id, cloneProfile(builtIn, builtIn.id, builtIn.name));
  }

  if (Array.isArray(loadedData.pdfProfiles)) {
    for (const rawProfile of loadedData.pdfProfiles) {
      if (rawProfile && typeof rawProfile.id === "string") {
        const validated = ensureValidProfile(rawProfile, rawProfile.id);
        profileMap.set(validated.id, validated);
      }
    }
  }

  settings.pdfProfiles = Array.from(profileMap.values());

  // Ensure default profile ID exists
  if (!settings.pdfProfiles.some((p) => p.id === settings.pdfDefaultProfileId)) {
    settings.pdfDefaultProfileId = ZOTERO_DEFAULT_PROFILE.id;
  }

  // 3. Per-PDF File Schemes
  if (loadedData.pdfFileSchemes && typeof loadedData.pdfFileSchemes === "object") {
    for (const [path, rawScheme] of Object.entries(loadedData.pdfFileSchemes)) {
      if (typeof path === "string" && rawScheme) {
        settings.pdfFileSchemes[path] = ensureValidProfile(rawScheme);
      }
    }
  }

  return settings;
}
