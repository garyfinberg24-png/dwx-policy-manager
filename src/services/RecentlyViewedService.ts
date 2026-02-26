/**
 * RecentlyViewedService — tracks recently viewed policies in localStorage.
 *
 * Storage key: pm_recently_viewed
 * Max items: 10 (most recent first)
 * Deduplicates on trackView (moves existing entry to top).
 *
 * Fully typed — no @ts-nocheck.
 */

/** Shape of a stored recently-viewed entry */
export interface IRecentlyViewedEntry {
  id: number;
  title: string;
  category: string;
  viewedAt: string; // ISO 8601 date string
}

/** Shape returned by getRecentlyViewed — adds a human-readable `time` field */
export interface IRecentlyViewedDisplay extends IRecentlyViewedEntry {
  time: string; // e.g. "5m ago", "2h ago", "1d ago"
}

const STORAGE_KEY = 'pm_recently_viewed';
const MAX_ITEMS = 10;

export class RecentlyViewedService {
  /**
   * Track a policy view. Deduplicates by id (moves existing to top) and
   * trims the list to MAX_ITEMS.
   */
  public static trackView(id: number, title: string, category: string): void {
    try {
      const entries = RecentlyViewedService.readEntries();

      // Remove any existing entry with the same id
      const filtered = entries.filter((e) => e.id !== id);

      // Prepend the new entry
      const newEntry: IRecentlyViewedEntry = {
        id,
        title,
        category,
        viewedAt: new Date().toISOString(),
      };
      filtered.unshift(newEntry);

      // Trim to max
      const trimmed = filtered.slice(0, MAX_ITEMS);

      localStorage.setItem(STORAGE_KEY, JSON.stringify(trimmed));
    } catch (err) {
      // localStorage may be unavailable (private browsing, quota exceeded, etc.)
      console.warn('RecentlyViewedService.trackView: unable to write localStorage', err);
    }
  }

  /**
   * Return the most recently viewed policies with a human-readable relative
   * time string.
   *
   * @param limit Maximum number of items to return (default 5).
   */
  public static getRecentlyViewed(limit: number = 5): IRecentlyViewedDisplay[] {
    try {
      const entries = RecentlyViewedService.readEntries();
      return entries.slice(0, limit).map((entry) => ({
        ...entry,
        time: RecentlyViewedService.formatRelativeTime(entry.viewedAt),
      }));
    } catch (err) {
      console.warn('RecentlyViewedService.getRecentlyViewed: unable to read localStorage', err);
      return [];
    }
  }

  /**
   * Clear all recently viewed history.
   */
  public static clearHistory(): void {
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch (err) {
      console.warn('RecentlyViewedService.clearHistory: unable to clear localStorage', err);
    }
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  /** Read and parse entries from localStorage. Returns [] on any error. */
  private static readEntries(): IRecentlyViewedEntry[] {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed: unknown = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      // Basic shape validation — keep entries that look correct
      return parsed.filter(
        (item: unknown): item is IRecentlyViewedEntry =>
          typeof item === 'object' &&
          item !== null &&
          typeof (item as IRecentlyViewedEntry).id === 'number' &&
          typeof (item as IRecentlyViewedEntry).title === 'string' &&
          typeof (item as IRecentlyViewedEntry).category === 'string' &&
          typeof (item as IRecentlyViewedEntry).viewedAt === 'string'
      );
    } catch {
      return [];
    }
  }

  /**
   * Convert an ISO date string to a short relative time label.
   * Examples: "just now", "5m ago", "2h ago", "1d ago", "3w ago"
   */
  private static formatRelativeTime(isoDate: string): string {
    const now = Date.now();
    const then = new Date(isoDate).getTime();
    const diffMs = now - then;

    if (diffMs < 0) return 'just now';

    const seconds = Math.floor(diffMs / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);
    const weeks = Math.floor(days / 7);

    if (seconds < 60) return 'just now';
    if (minutes < 60) return `${minutes}m ago`;
    if (hours < 24) return `${hours}h ago`;
    if (days < 7) return `${days}d ago`;
    if (weeks < 52) return `${weeks}w ago`;
    return `${Math.floor(weeks / 52)}y ago`;
  }
}
