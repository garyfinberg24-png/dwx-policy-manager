/**
 * RecentlyViewedService Unit Tests
 *
 * Tests trackView, getRecentlyViewed, deduplication, max items, clearHistory,
 * and relative time formatting.
 */

import {
  RecentlyViewedService,
  IRecentlyViewedEntry,
} from '../RecentlyViewedService';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const STORAGE_KEY = 'pm_recently_viewed';

/** Directly seed localStorage with entries for testing */
function seedStorage(entries: IRecentlyViewedEntry[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
}

/** Read raw entries from localStorage */
function readStorage(): IRecentlyViewedEntry[] {
  const raw = localStorage.getItem(STORAGE_KEY);
  return raw ? JSON.parse(raw) : [];
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('RecentlyViewedService', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  // ===== trackView =====

  describe('trackView', () => {
    it('should store a single entry in localStorage', () => {
      RecentlyViewedService.trackView(1, 'HR Policy', 'HR');
      const entries = readStorage();
      expect(entries).toHaveLength(1);
      expect(entries[0].id).toBe(1);
      expect(entries[0].title).toBe('HR Policy');
      expect(entries[0].category).toBe('HR');
      expect(entries[0].viewedAt).toBeDefined();
    });

    it('should prepend new entries (most recent first)', () => {
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      RecentlyViewedService.trackView(2, 'Policy B', 'HR');
      const entries = readStorage();
      expect(entries[0].id).toBe(2);
      expect(entries[1].id).toBe(1);
    });

    it('should deduplicate by id — move existing entry to top', () => {
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      RecentlyViewedService.trackView(2, 'Policy B', 'HR');
      RecentlyViewedService.trackView(1, 'Policy A Updated', 'IT');
      const entries = readStorage();
      expect(entries).toHaveLength(2);
      expect(entries[0].id).toBe(1);
      expect(entries[0].title).toBe('Policy A Updated');
      expect(entries[1].id).toBe(2);
    });

    it('should trim the list to 10 items max', () => {
      for (let i = 1; i <= 12; i++) {
        RecentlyViewedService.trackView(i, `Policy ${i}`, 'General');
      }
      const entries = readStorage();
      expect(entries).toHaveLength(10);
      // Most recent (12) should be first, oldest kept should be 3
      expect(entries[0].id).toBe(12);
      expect(entries[9].id).toBe(3);
    });

    it('should update viewedAt when re-viewing the same policy', () => {
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      const viewedAtBefore = readStorage()[0].viewedAt;
      expect(viewedAtBefore).toBeDefined();

      // Mock toISOString to return a different timestamp
      jest.spyOn(Date.prototype, 'toISOString').mockReturnValueOnce('2026-02-26T12:00:00.000Z');
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      const viewedAtAfter = readStorage()[0].viewedAt;

      expect(viewedAtAfter).toBe('2026-02-26T12:00:00.000Z');
    });
  });

  // ===== getRecentlyViewed =====

  describe('getRecentlyViewed', () => {
    it('should return an empty array when no history exists', () => {
      const result = RecentlyViewedService.getRecentlyViewed();
      expect(result).toEqual([]);
    });

    it('should return entries with a `time` field', () => {
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      const result = RecentlyViewedService.getRecentlyViewed();
      expect(result).toHaveLength(1);
      expect(result[0]).toHaveProperty('time');
      expect(result[0]).toHaveProperty('id', 1);
      expect(result[0]).toHaveProperty('title', 'Policy A');
    });

    it('should default to returning 5 items', () => {
      for (let i = 1; i <= 8; i++) {
        RecentlyViewedService.trackView(i, `Policy ${i}`, 'General');
      }
      const result = RecentlyViewedService.getRecentlyViewed();
      expect(result).toHaveLength(5);
    });

    it('should respect the limit parameter', () => {
      for (let i = 1; i <= 8; i++) {
        RecentlyViewedService.trackView(i, `Policy ${i}`, 'General');
      }
      expect(RecentlyViewedService.getRecentlyViewed(3)).toHaveLength(3);
      expect(RecentlyViewedService.getRecentlyViewed(10)).toHaveLength(8);
    });

    it('should return "just now" for very recent views', () => {
      RecentlyViewedService.trackView(1, 'Fresh', 'HR');
      const result = RecentlyViewedService.getRecentlyViewed(1);
      expect(result[0].time).toBe('just now');
    });

    it('should return relative time for older entries', () => {
      const twoHoursAgo = new Date(Date.now() - 2 * 60 * 60 * 1000).toISOString();
      seedStorage([
        { id: 1, title: 'Old Policy', category: 'HR', viewedAt: twoHoursAgo },
      ]);
      const result = RecentlyViewedService.getRecentlyViewed(1);
      expect(result[0].time).toBe('2h ago');
    });

    it('should return days for entries a few days old', () => {
      const threeDaysAgo = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000).toISOString();
      seedStorage([
        { id: 1, title: 'Old Policy', category: 'HR', viewedAt: threeDaysAgo },
      ]);
      const result = RecentlyViewedService.getRecentlyViewed(1);
      expect(result[0].time).toBe('3d ago');
    });

    it('should return weeks for entries older than a week', () => {
      const twoWeeksAgo = new Date(Date.now() - 14 * 24 * 60 * 60 * 1000).toISOString();
      seedStorage([
        { id: 1, title: 'Old Policy', category: 'HR', viewedAt: twoWeeksAgo },
      ]);
      const result = RecentlyViewedService.getRecentlyViewed(1);
      expect(result[0].time).toBe('2w ago');
    });
  });

  // ===== clearHistory =====

  describe('clearHistory', () => {
    it('should remove all entries from localStorage', () => {
      RecentlyViewedService.trackView(1, 'Policy A', 'IT');
      RecentlyViewedService.trackView(2, 'Policy B', 'HR');
      RecentlyViewedService.clearHistory();
      expect(localStorage.getItem(STORAGE_KEY)).toBeNull();
      expect(RecentlyViewedService.getRecentlyViewed()).toEqual([]);
    });

    it('should not throw when called with no history', () => {
      expect(() => RecentlyViewedService.clearHistory()).not.toThrow();
    });
  });

  // ===== Data validation =====

  describe('Data validation', () => {
    it('should ignore malformed entries in localStorage', () => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify([
        { id: 1, title: 'Good', category: 'HR', viewedAt: new Date().toISOString() },
        { broken: true }, // missing required fields
        'not an object',
        null,
        { id: 'not-a-number', title: 'Bad', category: 'HR', viewedAt: 'date' },
      ]));
      const result = RecentlyViewedService.getRecentlyViewed(10);
      expect(result).toHaveLength(1);
      expect(result[0].id).toBe(1);
    });

    it('should return empty array when localStorage has non-array JSON', () => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ not: 'an array' }));
      const result = RecentlyViewedService.getRecentlyViewed();
      expect(result).toEqual([]);
    });

    it('should return empty array when localStorage has invalid JSON', () => {
      localStorage.setItem(STORAGE_KEY, 'not valid json {{{');
      const result = RecentlyViewedService.getRecentlyViewed();
      expect(result).toEqual([]);
    });
  });
});
