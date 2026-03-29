/**
 * BookmarkService — tracks bookmarked policies in localStorage.
 *
 * Storage key: pm_bookmarked_policies
 * Max items: 25 (most recently bookmarked first)
 * Deduplicates on toggle (add/remove).
 *
 * NOTE: SP list persistence (PM_PolicyBookmarks) is planned for a future release.
 * Currently uses localStorage only — bookmarks are per-browser, not per-user.
 *
 * Fully typed — no @ts-nocheck.
 */

/** Shape of a stored bookmark entry */
export interface IBookmarkEntry {
  id: number;
  title: string;
  category: string;
  bookmarkedAt: string; // ISO 8601 date string
}

const STORAGE_KEY = 'pm_bookmarked_policies';
const MAX_ITEMS = 25;

export class BookmarkService {
  /**
   * Toggle a bookmark. If it exists, remove it; if not, add it.
   * Returns true if the policy is now bookmarked, false if removed.
   */
  public static toggle(id: number, title: string, category: string): boolean {
    try {
      const entries = BookmarkService.readEntries();
      const existingIndex = entries.findIndex((e) => e.id === id);

      if (existingIndex >= 0) {
        // Remove bookmark
        entries.splice(existingIndex, 1);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
        return false;
      } else {
        // Add bookmark at the top
        const newEntry: IBookmarkEntry = {
          id,
          title,
          category,
          bookmarkedAt: new Date().toISOString(),
        };
        entries.unshift(newEntry);
        const trimmed = entries.slice(0, MAX_ITEMS);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(trimmed));
        return true;
      }
    } catch (err) {
      console.warn('BookmarkService.toggle: unable to write localStorage', err);
      return false;
    }
  }

  /**
   * Check if a policy is bookmarked.
   */
  public static isBookmarked(id: number): boolean {
    try {
      const entries = BookmarkService.readEntries();
      return entries.some((e) => e.id === id);
    } catch {
      return false;
    }
  }

  /**
   * Return all bookmarked policies.
   */
  public static getAll(): IBookmarkEntry[] {
    try {
      return BookmarkService.readEntries();
    } catch {
      return [];
    }
  }

  /**
   * Return the set of bookmarked policy IDs (for quick lookup in components).
   */
  public static getBookmarkedIds(): Set<number> {
    try {
      const entries = BookmarkService.readEntries();
      return new Set(entries.map((e) => e.id));
    } catch {
      return new Set();
    }
  }

  /**
   * Remove a bookmark by policy ID.
   */
  public static remove(id: number): void {
    try {
      const entries = BookmarkService.readEntries().filter((e) => e.id !== id);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch (err) {
      console.warn('BookmarkService.remove: unable to write localStorage', err);
    }
  }

  /**
   * Clear all bookmarks.
   */
  public static clearAll(): void {
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {
      // localStorage may be unavailable
    }
  }

  /**
   * Read entries from localStorage with shape validation.
   */
  private static readEntries(): IBookmarkEntry[] {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];

      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];

      // Validate shape of each entry
      return parsed.filter(
        (item: unknown): item is IBookmarkEntry =>
          typeof item === 'object' &&
          item !== null &&
          typeof (item as IBookmarkEntry).id === 'number' &&
          typeof (item as IBookmarkEntry).title === 'string'
      );
    } catch {
      return [];
    }
  }
}
