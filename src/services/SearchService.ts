// @ts-nocheck
// Search Service
// Handles SharePoint Search API integration and advanced filtering

import { SPFI } from '@pnp/sp';
import '@pnp/sp/search';
import { ISearchQuery, ISearchResult, ISearchResponse, ISearchSuggestion, IRecentSearch, IAdvancedFilters } from '../models';
import { logger } from './LoggingService';

/**
 * Service for advanced search and filtering
 */
export class SearchService {
  private sp: SPFI;
  private readonly RECENT_SEARCHES_KEY = 'PM_recent_searches';
  private readonly MAX_RECENT_SEARCHES = 10;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Perform advanced search using SharePoint Search API
   */
  public async search(query: ISearchQuery): Promise<ISearchResponse> {
    try {
      const searchResults = await this.sp.search({
        Querytext: query.queryText,
        RowLimit: query.rowLimit || 50,
        StartRow: query.startRow || 0,
        SelectProperties: query.selectProperties || [
          'Title',
          'Path',
          'Description',
          'Created',
          'Modified',
          'Author',
          'EditorOWSUSER',
          'ProcessStatus',
          'ProcessType',
          'Department'
        ],
        EnableStemming: query.enableStemming !== false,
        EnablePhonetic: query.enablePhonetic || false,
        EnableNicknames: query.enableNicknames || false,
        TrimDuplicates: query.trimDuplicates !== false,
        SortList: query.sortList?.map(s => ({
          Property: s.property,
          Direction: s.direction === 'ascending' ? 0 : 1
        })),
        Refiners: query.refiners?.join(','),
        RefinementFilters: query.refinementFilters
      });

      const results: ISearchResult[] = [];
      if (searchResults.PrimarySearchResults) {
        for (let i = 0; i < searchResults.PrimarySearchResults.length; i++) {
          const item: any = searchResults.PrimarySearchResults[i];
          results.push({
            processId: this.extractProcessId(item.Path || ''),
            title: item.Title || '',
            description: item.Description || '',
            relevance: item.Rank || 0,
            highlights: this.extractHighlights(item.HitHighlightedSummary || ''),
            metadata: {
              path: item.Path,
              created: item.Created,
              modified: item.Modified,
              author: item.Author,
              status: item.ProcessStatus,
              type: item.ProcessType,
              department: item.Department
            }
          });
        }
      }

      const refinements = this.parseRefinements((searchResults as any).RefinementResults);

      return {
        results,
        totalResults: searchResults.TotalRows || 0,
        refinements,
        elapsedTime: searchResults.ElapsedTime
      };
    } catch (error) {
      logger.error('SearchService', 'Search failed:', error);
      throw error;
    }
  }

  /**
   * Build search query from advanced filters
   */
  public buildQueryFromFilters(filters: IAdvancedFilters): string {
    const queryParts: string[] = [];

    // Text search
    if (filters.searchText) {
      queryParts.push(`"${filters.searchText}"`);
    }

    // Status filter
    if (filters.statuses && filters.statuses.length > 0) {
      const statusQuery = filters.statuses.map(s => `ProcessStatus:"${s}"`).join(' OR ');
      queryParts.push(`(${statusQuery})`);
    }

    // Process type filter
    if (filters.processTypes && filters.processTypes.length > 0) {
      const typeQuery = filters.processTypes.map(t => `ProcessType:"${t}"`).join(' OR ');
      queryParts.push(`(${typeQuery})`);
    }

    // Department filter
    if (filters.departments && filters.departments.length > 0) {
      const deptQuery = filters.departments.map(d => `Department:"${d}"`).join(' OR ');
      queryParts.push(`(${deptQuery})`);
    }

    // Priority filter
    if (filters.priorities && filters.priorities.length > 0) {
      const priorityQuery = filters.priorities.map(p => `Priority:"${p}"`).join(' OR ');
      queryParts.push(`(${priorityQuery})`);
    }

    // Date range filters
    if (filters.createdDateRange) {
      const start = this.formatDateForSearch(filters.createdDateRange.start);
      const end = this.formatDateForSearch(filters.createdDateRange.end);
      queryParts.push(`Created>=${start} AND Created<=${end}`);
    }

    if (filters.dueDateRange) {
      const start = this.formatDateForSearch(filters.dueDateRange.start);
      const end = this.formatDateForSearch(filters.dueDateRange.end);
      queryParts.push(`DueDate>=${start} AND DueDate<=${end}`);
    }

    // People filters
    if (filters.managers && filters.managers.length > 0) {
      const managerQuery = filters.managers.map(m => `Manager:"${m.EMail}"`).join(' OR ');
      queryParts.push(`(${managerQuery})`);
    }

    if (filters.assignedTo && filters.assignedTo.length > 0) {
      const assignedQuery = filters.assignedTo.map(a => `AssignedTo:"${a.EMail}"`).join(' OR ');
      queryParts.push(`(${assignedQuery})`);
    }

    // Tags filter
    if (filters.tags && filters.tags.length > 0) {
      const tagQuery = filters.tags.map(t => `Tags:"${t}"`).join(' OR ');
      queryParts.push(`(${tagQuery})`);
    }

    // Boolean filters
    if (filters.hasOpenTasks === true) {
      queryParts.push('HasOpenTasks:true');
    }

    if (filters.isOverdue === true) {
      queryParts.push('IsOverdue:true');
    }

    if (filters.isCompleted === true) {
      queryParts.push('ProcessStatus:"Completed"');
    }

    // Combine all parts with AND
    return queryParts.length > 0 ? queryParts.join(' AND ') : '*';
  }

  /**
   * Get search suggestions for autocomplete
   */
  public async getSuggestions(queryText: string, maxSuggestions: number = 5): Promise<ISearchSuggestion[]> {
    try {
      const suggestions: ISearchSuggestion[] = [];

      // Get suggestions from SharePoint search
      const suggestResults = await this.sp.search({
        Querytext: `${queryText}*`,
        RowLimit: maxSuggestions,
        SelectProperties: ['Title', 'ProcessType', 'Department']
      });

      if (suggestResults.PrimarySearchResults) {
        for (let i = 0; i < suggestResults.PrimarySearchResults.length; i++) {
          const item: any = suggestResults.PrimarySearchResults[i];
          suggestions.push({
            text: item.Title || '',
            type: 'process',
            metadata: {
              type: item.ProcessType,
              department: item.Department
            }
          });
        }
      }

      // Add recent searches
      const recentSearches = this.getRecentSearches();
      for (let i = 0; i < recentSearches.length && suggestions.length < maxSuggestions; i++) {
        const recent = recentSearches[i];
        if (recent.searchText.toLowerCase().indexOf(queryText.toLowerCase()) !== -1) {
          suggestions.push({
            text: recent.searchText,
            type: 'recent',
            count: recent.resultCount
          });
        }
      }

      return suggestions;
    } catch (error) {
      logger.error('SearchService', 'Failed to get suggestions:', error);
      return [];
    }
  }

  /**
   * Save search to recent history
   */
  public saveRecentSearch(searchText: string, filters?: IAdvancedFilters, resultCount?: number): void {
    try {
      const recentSearches = this.getRecentSearches();

      const newSearch: IRecentSearch = {
        id: `search_${Date.now()}`,
        searchText,
        filters,
        timestamp: new Date(),
        resultCount
      };

      // Remove duplicate if exists - ES5 compatible
      const filtered = [];
      for (let i = 0; i < recentSearches.length; i++) {
        if (recentSearches[i].searchText !== searchText) {
          filtered.push(recentSearches[i]);
        }
      }

      // Add new search at beginning
      filtered.unshift(newSearch);

      // Keep only max recent searches
      const limited = filtered.slice(0, this.MAX_RECENT_SEARCHES);

      localStorage.setItem(this.RECENT_SEARCHES_KEY, JSON.stringify(limited));
    } catch (error) {
      logger.error('SearchService', 'Failed to save recent search:', error);
    }
  }

  /**
   * Get recent searches
   */
  public getRecentSearches(): IRecentSearch[] {
    try {
      const data = localStorage.getItem(this.RECENT_SEARCHES_KEY);
      if (!data) {
        return [];
      }

      const searches: IRecentSearch[] = JSON.parse(data);

      // Convert date strings to Date objects
      for (let i = 0; i < searches.length; i++) {
        if (typeof searches[i].timestamp === 'string') {
          searches[i].timestamp = new Date(searches[i].timestamp as any);
        }
      }

      return searches;
    } catch (error) {
      logger.error('SearchService', 'Failed to get recent searches:', error);
      return [];
    }
  }

  /**
   * Clear recent searches
   */
  public clearRecentSearches(): void {
    localStorage.removeItem(this.RECENT_SEARCHES_KEY);
  }

  /**
   * Search for people (for filters)
   */
  public async searchPeople(query: string): Promise<Array<{ id: number; name: string; email: string }>> {
    try {
      const results = await this.sp.search({
        Querytext: query,
        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31', // People source ID
        RowLimit: 10,
        SelectProperties: ['PreferredName', 'WorkEmail', 'AccountName']
      });

      const people = [];
      if (results.PrimarySearchResults) {
        for (let i = 0; i < results.PrimarySearchResults.length; i++) {
          const person: any = results.PrimarySearchResults[i];
          people.push({
            id: i,
            name: person.PreferredName || '',
            email: person.WorkEmail || person.AccountName || ''
          });
        }
      }

      return people;
    } catch (error) {
      logger.error('SearchService', 'Failed to search people:', error);
      return [];
    }
  }

  /**
   * Get available departments from search refiners
   */
  public async getDepartments(): Promise<string[]> {
    try {
      const results = await this.sp.search({
        Querytext: '*',
        RowLimit: 1,
        Refiners: 'Department',
        SelectProperties: ['Department']
      });

      const departments: string[] = [];
      const refinementResults = (results as any).RefinementResults;
      if (refinementResults) {
        for (let i = 0; i < refinementResults.length; i++) {
          const refiner = refinementResults[i];
          if (refiner.Name === 'Department' && refiner.Entries) {
            for (let j = 0; j < refiner.Entries.length; j++) {
              departments.push(refiner.Entries[j].RefinementName);
            }
          }
        }
      }

      return departments;
    } catch (error) {
      logger.error('SearchService', 'Failed to get departments:', error);
      return [];
    }
  }

  /**
   * Get available tags from search refiners
   */
  public async getTags(): Promise<string[]> {
    try {
      const results = await this.sp.search({
        Querytext: '*',
        RowLimit: 1,
        Refiners: 'Tags',
        SelectProperties: ['Tags']
      });

      const tags: string[] = [];
      const refinementResults = (results as any).RefinementResults;
      if (refinementResults) {
        for (let i = 0; i < refinementResults.length; i++) {
          const refiner = refinementResults[i];
          if (refiner.Name === 'Tags' && refiner.Entries) {
            for (let j = 0; j < refiner.Entries.length; j++) {
              tags.push(refiner.Entries[j].RefinementName);
            }
          }
        }
      }

      return tags;
    } catch (error) {
      logger.error('SearchService', 'Failed to get tags:', error);
      return [];
    }
  }

  // Private helper methods

  private extractProcessId(path: string): number {
    try {
      const match = path.match(/ID=(\d+)/);
      return match ? parseInt(match[1], 10) : 0;
    } catch {
      return 0;
    }
  }

  private extractHighlights(summary: string): Array<{ field: string; snippets: string[] }> {
    if (!summary) {
      return [];
    }

    return [{
      field: 'summary',
      snippets: [summary]
    }];
  }

  private parseRefinements(refinementResults: any): any[] {
    if (!refinementResults) {
      return [];
    }

    const refinements = [];
    for (let i = 0; i < refinementResults.length; i++) {
      const refiner = refinementResults[i];
      const entries = [];

      if (refiner.Entries) {
        for (let j = 0; j < refiner.Entries.length; j++) {
          const entry = refiner.Entries[j];
          entries.push({
            name: entry.RefinementName,
            count: entry.RefinementCount,
            token: entry.RefinementToken
          });
        }
      }

      refinements.push({
        name: refiner.Name,
        entries
      });
    }

    return refinements;
  }

  private formatDateForSearch(date: Date): string {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const monthStr = month < 10 ? '0' + month : String(month);
    const dayStr = day < 10 ? '0' + day : String(day);
    return `${year}-${monthStr}-${dayStr}`;
  }
}
