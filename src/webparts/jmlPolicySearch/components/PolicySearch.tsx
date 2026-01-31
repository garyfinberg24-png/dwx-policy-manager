// @ts-nocheck
import * as React from 'react';
import styles from './PolicySearch.module.scss';
import { IPolicySearchProps } from './IPolicySearchProps';
import {
  SearchBox,
  Icon,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Spinner,
  SpinnerSize,
  Checkbox,
  DefaultButton
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyService } from '../../../services/PolicyService';
import { SearchService } from '../../../services/SearchService';
import { IPolicy, PolicyStatus, PolicyCategory } from '../../../models/IPolicy';
import { logger } from '../../../services/LoggingService';

// Search result mapped from IPolicy
interface ISearchResult {
  id: number;
  type: 'policy';
  title: string;
  subtitle: string;
  description: string;
  status: string;
  category: string;
  lastModified: Date;
  policyNumber: string;
  versionNumber: string;
}

// Filter state interface
interface ISearchFilters {
  types: string[];
  status: string[];
  category: string[];
  dateFrom: Date | null;
  dateTo: Date | null;
}

interface IPolicySearchState {
  searchQuery: string;
  isSearching: boolean;
  results: ISearchResult[];
  totalResults: number;
  hasSearched: boolean;
  sortBy: string;
  filters: ISearchFilters;
  quickFilter: string | null;
  recentSearches: string[];
  isInitializing: boolean;
  error: string | null;
}

// Status options matching PolicyStatus enum values
const statusOptions = [
  { key: PolicyStatus.Draft, text: 'Draft' },
  { key: PolicyStatus.InReview, text: 'In Review' },
  { key: PolicyStatus.PendingApproval, text: 'Pending Approval' },
  { key: PolicyStatus.Published, text: 'Published' },
  { key: PolicyStatus.Approved, text: 'Approved' },
  { key: PolicyStatus.Archived, text: 'Archived' },
];

// Category options matching PolicyCategory enum values
const categoryOptions = [
  { key: PolicyCategory.HRPolicies, text: 'HR Policies' },
  { key: PolicyCategory.ITSecurity, text: 'IT & Security' },
  { key: PolicyCategory.HealthSafety, text: 'Health & Safety' },
  { key: PolicyCategory.Compliance, text: 'Compliance' },
  { key: PolicyCategory.Financial, text: 'Financial' },
  { key: PolicyCategory.Operational, text: 'Operational' },
  { key: PolicyCategory.Legal, text: 'Legal' },
  { key: PolicyCategory.Environmental, text: 'Environmental' },
  { key: PolicyCategory.QualityAssurance, text: 'Quality Assurance' },
  { key: PolicyCategory.DataPrivacy, text: 'Data Privacy' },
];

// Sort options
const sortOptions: IDropdownOption[] = [
  { key: 'PolicyName', text: 'Title A-Z' },
  { key: 'Modified', text: 'Most Recent' },
  { key: 'PolicyStatus', text: 'By Status' },
  { key: 'PolicyCategory', text: 'By Category' },
];

// Helper functions
const getStatusColor = (status: string): { bg: string; text: string } => {
  const colors: Record<string, { bg: string; text: string }> = {
    [PolicyStatus.Draft]: { bg: '#f3f2f1', text: '#605e5c' },
    [PolicyStatus.InReview]: { bg: '#e6f2ff', text: '#0078d4' },
    [PolicyStatus.PendingApproval]: { bg: '#fff4ce', text: '#8a6914' },
    [PolicyStatus.Approved]: { bg: '#dff6dd', text: '#107c10' },
    [PolicyStatus.Published]: { bg: '#dff6dd', text: '#107c10' },
    [PolicyStatus.Archived]: { bg: '#f3f2f1', text: '#605e5c' },
    [PolicyStatus.Retired]: { bg: '#f3f2f1', text: '#605e5c' },
    [PolicyStatus.Expired]: { bg: '#fde7e9', text: '#a4262c' },
    [PolicyStatus.Rejected]: { bg: '#fde7e9', text: '#a4262c' },
  };
  return colors[status] || { bg: '#f3f2f1', text: '#605e5c' };
};

const getCategoryIcon = (category: string): string => {
  const icons: Record<string, string> = {
    [PolicyCategory.HRPolicies]: 'People',
    [PolicyCategory.ITSecurity]: 'Shield',
    [PolicyCategory.HealthSafety]: 'Health',
    [PolicyCategory.Compliance]: 'ComplianceAudit',
    [PolicyCategory.Financial]: 'Money',
    [PolicyCategory.Operational]: 'Settings',
    [PolicyCategory.Legal]: 'Library',
    [PolicyCategory.Environmental]: 'Leaf',
    [PolicyCategory.QualityAssurance]: 'CheckMark',
    [PolicyCategory.DataPrivacy]: 'Lock',
  };
  return icons[category] || 'Library';
};

/**
 * Map an IPolicy to a search result for display
 */
function mapPolicyToResult(policy: IPolicy): ISearchResult {
  const modified = policy.Modified
    ? (policy.Modified instanceof Date ? policy.Modified : new Date(policy.Modified))
    : new Date();

  return {
    id: policy.Id || 0,
    type: 'policy',
    title: policy.PolicyName || policy.Title || '',
    subtitle: `${policy.PolicyNumber || ''} | ${policy.PolicyCategory || ''} | Version ${policy.VersionNumber || '1.0'}`,
    description: policy.Description || '',
    status: policy.PolicyStatus || PolicyStatus.Draft,
    category: policy.PolicyCategory || '',
    lastModified: modified,
    policyNumber: policy.PolicyNumber || '',
    versionNumber: policy.VersionNumber || '1.0',
  };
}

export default class PolicySearch extends React.Component<IPolicySearchProps, IPolicySearchState> {
  private policyService: PolicyService | null = null;
  private searchService: SearchService | null = null;

  constructor(props: IPolicySearchProps) {
    super(props);
    this.state = {
      searchQuery: '',
      isSearching: false,
      results: [],
      totalResults: 0,
      hasSearched: false,
      sortBy: 'Modified',
      filters: {
        types: [],
        status: [],
        category: [],
        dateFrom: null,
        dateTo: null,
      },
      quickFilter: null,
      recentSearches: [],
      isInitializing: true,
      error: null,
    };
  }

  public async componentDidMount(): Promise<void> {
    try {
      if (this.props.sp) {
        this.policyService = new PolicyService(this.props.sp);
        await this.policyService.initialize();

        this.searchService = new SearchService(this.props.sp);
        const recent = this.searchService.getRecentSearches();
        this.setState({
          recentSearches: recent.map(r => r.searchText).slice(0, 5),
          isInitializing: false,
        });

        // Check for search query in URL parameters (from header search bar)
        const urlParams = new URLSearchParams(window.location.search);
        const urlQuery = urlParams.get('q');
        if (urlQuery && urlQuery.trim()) {
          this.setState({ searchQuery: urlQuery.trim() }, () => {
            this.performSearch(urlQuery.trim());
          });
        }
      } else {
        this.setState({ isInitializing: false, error: 'SharePoint connection not available.' });
      }
    } catch (err) {
      logger.error('PolicySearch', 'Failed to initialize services', err);
      this.setState({ isInitializing: false, error: 'Failed to connect to SharePoint.' });
    }
  }

  private performSearch = async (query: string): Promise<void> => {
    if (!query.trim() && this.state.filters.status.length === 0 && this.state.filters.category.length === 0) {
      this.setState({ results: [], totalResults: 0, hasSearched: false });
      return;
    }

    if (!this.policyService) {
      this.setState({ error: 'Policy service not available.' });
      return;
    }

    this.setState({ isSearching: true, hasSearched: true, error: null });

    try {
      const { filters, sortBy } = this.state;

      // Determine sort direction
      const sortDirection: 'asc' | 'desc' = sortBy === 'Modified' ? 'desc' : 'asc';

      // Build status filter — only use first selected if PolicyService expects single value
      const statusFilter = filters.status.length === 1
        ? filters.status[0] as PolicyStatus
        : undefined;

      // Build category filter — only use first selected if PolicyService expects single value
      const categoryFilter = filters.category.length === 1
        ? filters.category[0]
        : undefined;

      // Fetch policies from SharePoint
      const result = await this.policyService.getPoliciesPaginated(
        1,
        50,
        {
          searchTerm: query.trim() || undefined,
          status: statusFilter,
          category: categoryFilter,
          sortBy: sortBy,
          sortDirection: sortDirection,
        }
      );

      let searchResults = result.items.map(mapPolicyToResult);

      // Apply client-side filters for multi-select status/category (service supports single only)
      if (filters.status.length > 1) {
        searchResults = searchResults.filter(r => filters.status.includes(r.status));
      }
      if (filters.category.length > 1) {
        searchResults = searchResults.filter(r => filters.category.includes(r.category));
      }

      // Apply date filters client-side
      if (filters.dateFrom) {
        const from = filters.dateFrom.getTime();
        searchResults = searchResults.filter(r => r.lastModified.getTime() >= from);
      }
      if (filters.dateTo) {
        const to = filters.dateTo.getTime();
        searchResults = searchResults.filter(r => r.lastModified.getTime() <= to);
      }

      this.setState({
        results: searchResults,
        totalResults: searchResults.length,
        isSearching: false,
      });

      // Save to recent searches
      if (query.trim() && this.searchService) {
        this.searchService.saveRecentSearch(query.trim(), undefined, searchResults.length);
        const recent = this.searchService.getRecentSearches();
        this.setState({ recentSearches: recent.map(r => r.searchText).slice(0, 5) });
      }
    } catch (err) {
      logger.error('PolicySearch', 'Search failed', err);
      this.setState({
        isSearching: false,
        error: 'Search failed. Please try again.',
        results: [],
        totalResults: 0,
      });
    }
  };

  private handleSearch = (): void => {
    this.performSearch(this.state.searchQuery);
  };

  private handleFilterChange = (filterType: string, value: string, checked: boolean): void => {
    this.setState(prevState => {
      const currentValues = (prevState.filters as any)[filterType] as string[];
      const newValues = checked
        ? [...currentValues, value]
        : currentValues.filter((v: string) => v !== value);

      return {
        filters: { ...prevState.filters, [filterType]: newValues }
      };
    }, () => {
      if (this.state.hasSearched) {
        this.performSearch(this.state.searchQuery);
      }
    });
  };

  private clearFilters = (): void => {
    this.setState({
      filters: {
        types: [],
        status: [],
        category: [],
        dateFrom: null,
        dateTo: null,
      },
      quickFilter: null,
    }, () => {
      if (this.state.hasSearched) {
        this.performSearch(this.state.searchQuery);
      }
    });
  };

  private handleQuickFilter = (category: string): void => {
    const { quickFilter } = this.state;
    if (quickFilter === category) {
      this.setState(prevState => ({
        quickFilter: null,
        filters: { ...prevState.filters, category: [] }
      }), () => {
        if (this.state.hasSearched) this.performSearch(this.state.searchQuery);
      });
    } else {
      this.setState(prevState => ({
        quickFilter: category,
        filters: { ...prevState.filters, category: [category] }
      }), () => {
        // Auto-search when quick filter is applied
        this.performSearch(this.state.searchQuery || '*');
      });
    }
  };

  private handleResultClick = (result: ISearchResult): void => {
    if (result.type === 'policy' && result.id) {
      const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
      window.location.href = `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${result.id}&mode=browse`;
    }
  };

  public render(): React.ReactElement<IPolicySearchProps> {
    const { searchQuery, isSearching, results, totalResults, hasSearched, sortBy, filters, quickFilter, recentSearches, isInitializing, error } = this.state;
    const activeFilterCount = filters.status.length + filters.category.length +
      (filters.dateFrom ? 1 : 0) + (filters.dateTo ? 1 : 0);

    // Quick filter chips — use top categories
    const quickFilterOptions = [
      { key: PolicyCategory.HRPolicies, text: 'HR Policies', icon: 'People' },
      { key: PolicyCategory.ITSecurity, text: 'IT & Security', icon: 'Shield' },
      { key: PolicyCategory.Compliance, text: 'Compliance', icon: 'ComplianceAudit' },
      { key: PolicyCategory.HealthSafety, text: 'Health & Safety', icon: 'Health' },
      { key: PolicyCategory.Financial, text: 'Financial', icon: 'Money' },
    ];

    return (
      <JmlAppLayout context={this.props.context} breadcrumbs={[{ text: 'Policy Manager', url: this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager' }, { text: 'Search' }]}>
        <div className={styles.policySearch}>
          <div className={styles.contentWrapper}>
            {/* Hero Section */}
            <div className={styles.heroSection}>
              <div className={styles.heroHeader}>
                <Icon iconName="Search" className={styles.heroIcon} />
                <div>
                  <h1 className={styles.heroTitle}>Search Policies</h1>
                  <p className={styles.heroSubtitle}>
                    Find policies by name, number, keywords, or category
                  </p>
                </div>
              </div>

              {isInitializing ? (
                <Spinner size={SpinnerSize.medium} label="Connecting to SharePoint..." />
              ) : (
                <>
                  <div className={styles.searchBoxContainer}>
                    <div className={styles.searchInputRow}>
                      <SearchBox
                        placeholder="Search by policy name, number, keywords, category..."
                        value={searchQuery}
                        onChange={(_, value) => this.setState({ searchQuery: value || '' })}
                        onSearch={this.handleSearch}
                        onClear={() => this.setState({ searchQuery: '', results: [], totalResults: 0, hasSearched: false, error: null })}
                        styles={{
                          root: { borderRadius: '6px', backgroundColor: '#ffffff', flex: 1 },
                          field: { fontSize: '16px', padding: '8px' },
                        }}
                      />
                      <button
                        className={styles.searchButton}
                        onClick={() => this.handleSearch(searchQuery)}
                        type="button"
                        title="Search"
                      >
                        <Icon iconName="Search" />
                        Search
                      </button>
                    </div>

                    {/* Quick Filters by Category */}
                    <div className={styles.quickFilters}>
                      {quickFilterOptions.map(opt => (
                        <button
                          key={opt.key}
                          className={`${styles.quickFilterChip} ${quickFilter === opt.key ? styles.quickFilterChipActive : ''}`}
                          onClick={() => this.handleQuickFilter(opt.key)}
                          type="button"
                        >
                          <Icon iconName={opt.icon} />
                          {opt.text}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Error Message */}
                  {error && (
                    <div style={{ color: '#a4262c', padding: '8px 16px', backgroundColor: '#fde7e9', borderRadius: '4px', marginTop: '8px' }}>
                      <Icon iconName="ErrorBadge" style={{ marginRight: '8px' }} />
                      {error}
                    </div>
                  )}

                  {/* Recent Searches */}
                  {!hasSearched && recentSearches.length > 0 && (
                    <div className={styles.recentSearches}>
                      <div className={styles.recentSearchTitle}>Recent Searches</div>
                      <div className={styles.recentSearchTags}>
                        {recentSearches.map((search, index) => (
                          <span
                            key={index}
                            className={styles.recentSearchTag}
                            onClick={() => {
                              this.setState({ searchQuery: search }, () => {
                                this.performSearch(search);
                              });
                            }}
                          >
                            <Icon iconName="History" />
                            {search}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>

            {/* Search Results Area */}
            {hasSearched && (
              <div className={styles.mainContent}>
                {/* Filters Panel */}
                <div className={styles.filtersPanel}>
                  <h3 className={styles.filtersPanelTitle}>
                    Filters
                    {activeFilterCount > 0 && (
                      <span className={styles.filterCount}>({activeFilterCount})</span>
                    )}
                  </h3>

                  {/* Status Filter */}
                  <div className={styles.filterSection}>
                    <div className={styles.filterTitle}>Status</div>
                    {statusOptions.map(opt => (
                      <Checkbox
                        key={opt.key}
                        label={opt.text}
                        checked={filters.status.includes(opt.key)}
                        onChange={(_, checked) => this.handleFilterChange('status', opt.key, checked || false)}
                        className={styles.filterCheckbox}
                      />
                    ))}
                  </div>

                  {/* Category Filter */}
                  <div className={styles.filterSection}>
                    <div className={styles.filterTitle}>Category</div>
                    {categoryOptions.map(opt => (
                      <Checkbox
                        key={opt.key}
                        label={opt.text}
                        checked={filters.category.includes(opt.key)}
                        onChange={(_, checked) => this.handleFilterChange('category', opt.key, checked || false)}
                        className={styles.filterCheckbox}
                      />
                    ))}
                  </div>

                  {/* Date Filter */}
                  <div className={styles.filterSection}>
                    <div className={styles.filterTitle}>Last Modified</div>
                    <DatePicker
                      placeholder="From date"
                      value={filters.dateFrom || undefined}
                      onSelectDate={(date) => this.setState(prev => ({
                        filters: { ...prev.filters, dateFrom: date || null }
                      }), () => {
                        if (this.state.hasSearched) this.performSearch(this.state.searchQuery);
                      })}
                      styles={{ root: { marginBottom: '8px' } }}
                    />
                    <DatePicker
                      placeholder="To date"
                      value={filters.dateTo || undefined}
                      onSelectDate={(date) => this.setState(prev => ({
                        filters: { ...prev.filters, dateTo: date || null }
                      }), () => {
                        if (this.state.hasSearched) this.performSearch(this.state.searchQuery);
                      })}
                    />
                  </div>

                  {activeFilterCount > 0 && (
                    <button className={styles.clearFilters} onClick={this.clearFilters} type="button">
                      <Icon iconName="Cancel" />
                      Clear all filters
                    </button>
                  )}
                </div>

                {/* Results Panel */}
                <div className={styles.resultsPanel}>
                  {isSearching ? (
                    <div className={styles.loadingOverlay}>
                      <Spinner size={SpinnerSize.large} label="Searching policies..." />
                    </div>
                  ) : results.length > 0 ? (
                    <>
                      <div className={styles.resultsHeader}>
                        <span className={styles.resultsCount}>
                          {totalResults} {totalResults === 1 ? 'result' : 'results'}
                          {searchQuery.trim() && <> for &ldquo;{searchQuery}&rdquo;</>}
                        </span>
                        <Dropdown
                          placeholder="Sort by"
                          selectedKey={sortBy}
                          options={sortOptions}
                          onChange={(_, option) => {
                            this.setState({ sortBy: (option?.key as string) || 'Modified' }, () => {
                              this.performSearch(this.state.searchQuery);
                            });
                          }}
                          className={styles.sortDropdown}
                        />
                      </div>

                      {results.map(result => (
                        <div
                          key={`policy-${result.id}`}
                          className={styles.resultCard}
                          onClick={() => this.handleResultClick(result)}
                        >
                          <div className={styles.resultHeader}>
                            <div
                              className={styles.resultIcon}
                              style={{ backgroundColor: '#0d9488' }}
                            >
                              <Icon iconName={getCategoryIcon(result.category)} />
                            </div>
                            <div className={styles.resultInfo}>
                              <div className={styles.resultTitle}>{result.title}</div>
                              <div className={styles.resultSubtitle}>{result.subtitle}</div>
                            </div>
                            {result.status && (
                              <span
                                className={styles.statusBadge}
                                style={{
                                  backgroundColor: getStatusColor(result.status).bg,
                                  color: getStatusColor(result.status).text,
                                }}
                              >
                                {result.status}
                              </span>
                            )}
                          </div>

                          {result.description && (
                            <div className={styles.resultHighlights}>
                              {result.description.length > 200
                                ? result.description.substring(0, 200) + '...'
                                : result.description}
                            </div>
                          )}

                          <div className={styles.resultMeta}>
                            <span className={styles.resultMetaItem}>
                              <Icon iconName={getCategoryIcon(result.category)} />
                              {result.category}
                            </span>
                            <span className={styles.resultMetaItem}>
                              <Icon iconName="Calendar" />
                              Modified {result.lastModified.toLocaleDateString()}
                            </span>
                            {result.policyNumber && (
                              <span className={styles.resultMetaItem}>
                                <Icon iconName="NumberField" />
                                {result.policyNumber}
                              </span>
                            )}
                          </div>
                        </div>
                      ))}
                    </>
                  ) : (
                    <div className={styles.noResults}>
                      <Icon iconName="SearchIssue" className={styles.noResultsIcon} />
                      <div className={styles.noResultsTitle}>No results found</div>
                      <div className={styles.noResultsText}>
                        {searchQuery.trim()
                          ? <>We couldn&apos;t find any policies matching &ldquo;{searchQuery}&rdquo;. Try different keywords or adjust your filters.</>
                          : <>No policies match the selected filters. Try adjusting your filter criteria.</>
                        }
                      </div>
                      {activeFilterCount > 0 && (
                        <DefaultButton
                          text="Clear Filters"
                          onClick={this.clearFilters}
                          style={{ marginTop: '16px' }}
                        />
                      )}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      </JmlAppLayout>
    );
  }
}
