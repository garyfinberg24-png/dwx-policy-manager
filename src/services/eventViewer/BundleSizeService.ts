/**
 * BundleSizeService — Analyses JS bundle sizes for each webpart
 * by reading performance.getEntriesByType('resource').
 * Provides a breakdown of loaded JS/CSS resources by webpart.
 */

// ============================================================================
// TYPES
// ============================================================================

export interface IBundleEntry {
  name: string;
  /** Shortened URL for display */
  shortName: string;
  /** Transfer size in bytes */
  transferSize: number;
  /** Decoded body size in bytes */
  decodedSize: number;
  /** Resource type (script, stylesheet, other) */
  type: 'script' | 'stylesheet' | 'other';
  /** Duration to load in ms */
  duration: number;
  /** Webpart name if identifiable from URL */
  webpart?: string;
}

export interface IBundleSizeSummary {
  entries: IBundleEntry[];
  totalTransferKB: number;
  totalDecodedKB: number;
  scriptCount: number;
  styleCount: number;
  webpartBreakdown: Array<{ name: string; sizeKB: number; fileCount: number }>;
}

// ============================================================================
// WEBPART PATTERNS — map URL segments to webpart names
// ============================================================================

const WEBPART_PATTERNS: Array<{ pattern: RegExp; name: string }> = [
  { pattern: /jml-my-policies/i, name: 'MyPolicies' },
  { pattern: /jml-policy-hub/i, name: 'PolicyHub' },
  { pattern: /jml-policy-admin/i, name: 'PolicyAdmin' },
  { pattern: /jml-policy-author/i, name: 'PolicyAuthor' },
  { pattern: /jml-policy-details/i, name: 'PolicyDetails' },
  { pattern: /jml-policy-pack/i, name: 'PolicyPacks' },
  { pattern: /dwx-quiz-builder/i, name: 'QuizBuilder' },
  { pattern: /jml-policy-search/i, name: 'PolicySearch' },
  { pattern: /jml-policy-help/i, name: 'PolicyHelp' },
  { pattern: /jml-policy-distribution/i, name: 'PolicyDistribution' },
  { pattern: /jml-policy-analytics/i, name: 'PolicyAnalytics' },
  { pattern: /dwx-policy-author-view/i, name: 'AuthorView' },
  { pattern: /dwx-policy-manager-view/i, name: 'ManagerView' },
  { pattern: /dwx-policy-author-reports/i, name: 'AuthorReports' },
  { pattern: /dwx-policy-bulk-upload/i, name: 'BulkUpload' },
  { pattern: /dwx-event-viewer/i, name: 'EventViewer' },
  { pattern: /vendor|chunk|common|polyfill/i, name: 'Shared/Vendor' },
];

// ============================================================================
// SERVICE
// ============================================================================

export class BundleSizeService {

  /**
   * Analyse loaded resources using the Performance API.
   */
  public static analyse(): IBundleSizeSummary {
    const entries: IBundleEntry[] = [];

    if (typeof performance === 'undefined' || !performance.getEntriesByType) {
      return { entries: [], totalTransferKB: 0, totalDecodedKB: 0, scriptCount: 0, styleCount: 0, webpartBreakdown: [] };
    }

    const resources = performance.getEntriesByType('resource') as PerformanceResourceTiming[];

    for (const r of resources) {
      const isScript = r.initiatorType === 'script' || r.name.endsWith('.js');
      const isStyle = r.initiatorType === 'link' || r.name.endsWith('.css');
      if (!isScript && !isStyle) continue;

      const type: IBundleEntry['type'] = isScript ? 'script' : isStyle ? 'stylesheet' : 'other';
      const shortName = BundleSizeService._shortenUrl(r.name);
      const webpart = BundleSizeService._identifyWebpart(r.name);

      entries.push({
        name: r.name,
        shortName,
        transferSize: r.transferSize || 0,
        decodedSize: r.decodedBodySize || 0,
        type,
        duration: Math.round(r.duration),
        webpart,
      });
    }

    // Sort by transfer size descending
    entries.sort((a, b) => b.transferSize - a.transferSize);

    const totalTransfer = entries.reduce((s, e) => s + e.transferSize, 0);
    const totalDecoded = entries.reduce((s, e) => s + e.decodedSize, 0);

    // Webpart breakdown
    const wpMap: Record<string, { sizeKB: number; fileCount: number }> = {};
    for (const e of entries) {
      const wp = e.webpart || 'Other';
      if (!wpMap[wp]) wpMap[wp] = { sizeKB: 0, fileCount: 0 };
      wpMap[wp].sizeKB += e.transferSize / 1024;
      wpMap[wp].fileCount++;
    }

    const webpartBreakdown = Object.entries(wpMap)
      .map(([name, data]) => ({ name, sizeKB: Math.round(data.sizeKB * 10) / 10, fileCount: data.fileCount }))
      .sort((a, b) => b.sizeKB - a.sizeKB);

    return {
      entries,
      totalTransferKB: Math.round(totalTransfer / 1024 * 10) / 10,
      totalDecodedKB: Math.round(totalDecoded / 1024 * 10) / 10,
      scriptCount: entries.filter(e => e.type === 'script').length,
      styleCount: entries.filter(e => e.type === 'stylesheet').length,
      webpartBreakdown,
    };
  }

  private static _shortenUrl(url: string): string {
    const idx = url.lastIndexOf('/');
    if (idx !== -1) {
      const filename = url.substring(idx + 1);
      const qIdx = filename.indexOf('?');
      return qIdx !== -1 ? filename.substring(0, qIdx) : filename;
    }
    return url.length > 60 ? url.substring(url.length - 57) + '...' : url;
  }

  private static _identifyWebpart(url: string): string | undefined {
    for (const p of WEBPART_PATTERNS) {
      if (p.pattern.test(url)) return p.name;
    }
    return undefined;
  }
}
