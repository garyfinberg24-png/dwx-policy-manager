// Advanced Search and Filter Models
// Interfaces for advanced search, filtering, and user preferences

import { IUser } from './ICommon';
import { IJmlTagSelection, JmlTermSetType } from './ITaxonomy';

// Person type alias for consistency with search models
export type IPerson = IUser;

/**
 * Advanced filter options for process searching
 */
export interface IAdvancedFilters {
  // Text search
  searchText?: string;

  // Date filters
  dateRange?: {
    start: Date;
    end: Date;
  };
  createdDateRange?: {
    start: Date;
    end: Date;
  };
  dueDateRange?: {
    start: Date;
    end: Date;
  };
  completedDateRange?: {
    start: Date;
    end: Date;
  };

  // Categorical filters
  departments?: string[];
  processTypes?: string[];
  statuses?: string[];
  priorities?: string[];

  // People filters
  managers?: IPerson[];
  assignedTo?: IPerson[];
  createdBy?: IPerson[];

  // Numeric filters
  costRange?: {
    min: number;
    max: number;
  };
  progressRange?: {
    min: number;
    max: number;
  };

  // Tag and metadata filters (legacy string-based)
  tags?: string[];

  // Taxonomy-based tag filters (Managed Metadata)
  taxonomyTags?: IJmlTagSelection[];
  taxonomyFilters?: {
    termSetType: JmlTermSetType;
    termIds: string[];
  }[];

  customFields?: { [key: string]: any };

  // Boolean filters
  hasOpenTasks?: boolean;
  isOverdue?: boolean;
  isCompleted?: boolean;
  hasAttachments?: boolean;
}

/**
 * Saved filter preset for quick access
 */
export interface IFilterPreset {
  id: string;
  title: string;
  description?: string;
  filters: IAdvancedFilters;
  userId: number;
  isDefault?: boolean;
  isShared?: boolean;
  createdDate: Date;
  modifiedDate: Date;
  useCount?: number;
}

/**
 * Search suggestion item
 */
export interface ISearchSuggestion {
  text: string;
  type: 'process' | 'person' | 'department' | 'tag' | 'recent';
  count?: number;
  metadata?: any;
}

/**
 * Recent search history item
 */
export interface IRecentSearch {
  id: string;
  searchText: string;
  filters?: IAdvancedFilters;
  timestamp: Date;
  resultCount?: number;
}

/**
 * Search result item with highlighting
 */
export interface ISearchResult {
  processId: number;
  title: string;
  description: string;
  relevance: number;
  highlights?: {
    field: string;
    snippets: string[];
  }[];
  metadata?: any;
}

/**
 * SharePoint Search query configuration
 */
export interface ISearchQuery {
  queryText: string;
  rowLimit?: number;
  startRow?: number;
  sortList?: Array<{
    property: string;
    direction: 'ascending' | 'descending';
  }>;
  selectProperties?: string[];
  refiners?: string[];
  refinementFilters?: string[];
  enableStemming?: boolean;
  enablePhonetic?: boolean;
  enableNicknames?: boolean;
  trimDuplicates?: boolean;
}

/**
 * Search refinement result
 */
export interface ISearchRefinement {
  name: string;
  entries: Array<{
    name: string;
    count: number;
    token: string;
  }>;
}

/**
 * Complete search response
 */
export interface ISearchResponse {
  results: ISearchResult[];
  totalResults: number;
  refinements?: ISearchRefinement[];
  elapsedTime?: number;
}

/**
 * Filter options for autocomplete
 */
export interface IAutocompleteOptions {
  minLength?: number;
  maxSuggestions?: number;
  debounceMs?: number;
  includeRecent?: boolean;
  includeSuggestions?: boolean;
}

/**
 * User search preferences
 */
export interface ISearchPreferences {
  userId: number;
  defaultFilters?: IAdvancedFilters;
  recentSearches: IRecentSearch[];
  maxRecentSearches?: number;
  savedPresets: string[]; // IDs of saved filter presets
  searchHistory?: {
    query: string;
    timestamp: Date;
  }[];
}
