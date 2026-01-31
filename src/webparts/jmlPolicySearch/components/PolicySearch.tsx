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

// Search result interface
interface ISearchResult {
  id: number;
  type: 'policy' | 'template' | 'pack' | 'document';
  title: string;
  subtitle: string;
  highlights: string[];
  status?: string;
  category?: string;
  lastModified: Date;
  relevanceScore: number;
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
  hasSearched: boolean;
  sortBy: string;
  filters: ISearchFilters;
  quickFilter: string | null;
  recentSearches: string[];
}

// Type options for policy search
const typeOptions = [
  { key: 'policy', text: 'Policies', icon: 'Library' },
  { key: 'template', text: 'Templates', icon: 'FileTemplate' },
  { key: 'pack', text: 'Policy Packs', icon: 'Package' },
  { key: 'document', text: 'Documents', icon: 'Document' },
];

// Status options
const statusOptions = [
  { key: 'draft', text: 'Draft' },
  { key: 'pending', text: 'Pending Approval' },
  { key: 'published', text: 'Published' },
  { key: 'review', text: 'Under Review' },
  { key: 'archived', text: 'Archived' },
];

// Category options
const categoryOptions = [
  { key: 'hr', text: 'Human Resources' },
  { key: 'it', text: 'Information Technology' },
  { key: 'compliance', text: 'Compliance & Legal' },
  { key: 'finance', text: 'Finance' },
  { key: 'operations', text: 'Operations' },
  { key: 'safety', text: 'Health & Safety' },
  { key: 'security', text: 'Security' },
  { key: 'governance', text: 'Governance' },
];

// Sort options
const sortOptions: IDropdownOption[] = [
  { key: 'relevance', text: 'Most Relevant' },
  { key: 'recent', text: 'Most Recent' },
  { key: 'title', text: 'Title A-Z' },
  { key: 'status', text: 'By Status' },
];

// Mock search results generator
const generateMockResults = (query: string, filters: ISearchFilters): ISearchResult[] => {
  if (!query.trim()) return [];

  const mockData: ISearchResult[] = [
    {
      id: 1,
      type: 'policy',
      title: 'Information Security Policy',
      subtitle: 'POL-IT-001 | IT Department | Version 3.2',
      highlights: [`Comprehensive information security guidelines covering "${query}" requirements and data protection standards...`],
      status: 'published',
      category: 'it',
      lastModified: new Date('2024-01-15'),
      relevanceScore: 0.95,
    },
    {
      id: 2,
      type: 'policy',
      title: 'Data Privacy & Protection Policy',
      subtitle: 'POL-COMP-003 | Compliance | Version 2.1',
      highlights: [`Data handling procedures relating to "${query}" and GDPR compliance requirements...`],
      status: 'published',
      category: 'compliance',
      lastModified: new Date('2024-01-10'),
      relevanceScore: 0.88,
    },
    {
      id: 3,
      type: 'template',
      title: 'Standard Policy Template v2.0',
      subtitle: 'Template | 8 sections | Governance category',
      highlights: [`Reusable policy template with sections covering "${query}" and related governance topics...`],
      status: undefined,
      category: 'governance',
      lastModified: new Date('2023-12-01'),
      relevanceScore: 0.82,
    },
    {
      id: 4,
      type: 'pack',
      title: 'New Employee Onboarding Pack',
      subtitle: 'Policy Pack | 12 policies | HR Department',
      highlights: [`Employee onboarding bundle including policies related to "${query}" and workplace conduct...`],
      status: 'published',
      category: 'hr',
      lastModified: new Date('2024-01-05'),
      relevanceScore: 0.75,
    },
    {
      id: 5,
      type: 'policy',
      title: 'Acceptable Use Policy',
      subtitle: 'POL-IT-005 | IT Department | Version 1.8',
      highlights: [`Guidelines for acceptable use of company resources relating to "${query}" and technology usage...`],
      status: 'review',
      category: 'it',
      lastModified: new Date('2024-01-14'),
      relevanceScore: 0.70,
    },
    {
      id: 6,
      type: 'document',
      title: 'Compliance Audit Report Q4 2023',
      subtitle: 'Document | Compliance Department | 45 pages',
      highlights: [`Quarterly compliance audit findings referencing "${query}" and regulatory adherence...`],
      status: undefined,
      category: 'compliance',
      lastModified: new Date('2023-11-20'),
      relevanceScore: 0.65,
    },
  ];

  let results = mockData;
  if (filters.types.length > 0) {
    results = results.filter(r => filters.types.includes(r.type));
  }
  if (filters.status.length > 0) {
    results = results.filter(r => !r.status || filters.status.includes(r.status));
  }
  if (filters.category.length > 0) {
    results = results.filter(r => !r.category || filters.category.includes(r.category));
  }
  return results;
};

// Helper functions
const getTypeColor = (type: string): string => {
  const colors: Record<string, string> = {
    policy: '#0d9488',
    template: '#8764b8',
    pack: '#107c10',
    document: '#0078d4',
  };
  return colors[type] || '#0d9488';
};

const getTypeIcon = (type: string): string => {
  const icons: Record<string, string> = {
    policy: 'Library',
    template: 'FileTemplate',
    pack: 'Package',
    document: 'Document',
  };
  return icons[type] || 'Document';
};

const getStatusColor = (status: string): { bg: string; text: string } => {
  const colors: Record<string, { bg: string; text: string }> = {
    draft: { bg: '#f3f2f1', text: '#605e5c' },
    pending: { bg: '#fff4ce', text: '#8a6914' },
    published: { bg: '#dff6dd', text: '#107c10' },
    review: { bg: '#e6f2ff', text: '#0078d4' },
    archived: { bg: '#f3f2f1', text: '#605e5c' },
  };
  return colors[status] || { bg: '#f3f2f1', text: '#605e5c' };
};

export default class PolicySearch extends React.Component<IPolicySearchProps, IPolicySearchState> {
  constructor(props: IPolicySearchProps) {
    super(props);
    this.state = {
      searchQuery: '',
      isSearching: false,
      results: [],
      hasSearched: false,
      sortBy: 'relevance',
      filters: {
        types: [],
        status: [],
        category: [],
        dateFrom: null,
        dateTo: null,
      },
      quickFilter: null,
      recentSearches: [
        'Data Protection',
        'Information Security',
        'Code of Conduct',
        'Leave Policy',
        'Compliance',
      ],
    };
  }

  private performSearch = async (query: string): Promise<void> => {
    if (!query.trim()) {
      this.setState({ results: [], hasSearched: false });
      return;
    }

    this.setState({ isSearching: true, hasSearched: true });

    // Simulate API call
    await new Promise(resolve => setTimeout(resolve, 800));

    const searchResults = generateMockResults(query, this.state.filters);
    this.setState({ results: searchResults, isSearching: false });
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
      if (this.state.hasSearched && this.state.searchQuery) {
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
      if (this.state.hasSearched && this.state.searchQuery) {
        this.performSearch(this.state.searchQuery);
      }
    });
  };

  private handleQuickFilter = (type: string): void => {
    const { quickFilter } = this.state;
    if (quickFilter === type) {
      this.setState(prevState => ({
        quickFilter: null,
        filters: { ...prevState.filters, types: [] }
      }), () => {
        if (this.state.hasSearched) this.performSearch(this.state.searchQuery);
      });
    } else {
      this.setState(prevState => ({
        quickFilter: type,
        filters: { ...prevState.filters, types: [type] }
      }), () => {
        if (this.state.hasSearched) this.performSearch(this.state.searchQuery);
      });
    }
  };

  private handleResultClick = (result: ISearchResult): void => {
    if (result.type === 'policy') {
      window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${result.id}&mode=browse`;
    }
  };

  public render(): React.ReactElement<IPolicySearchProps> {
    const { searchQuery, isSearching, results, hasSearched, sortBy, filters, quickFilter, recentSearches } = this.state;
    const activeFilterCount = filters.types.length + filters.status.length + filters.category.length +
      (filters.dateFrom ? 1 : 0) + (filters.dateTo ? 1 : 0);

    return (
      <JmlAppLayout context={this.props.context} breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Search' }]}>
        <div className={styles.policySearch}>
          <div className={styles.contentWrapper}>
            {/* Hero Section */}
            <div className={styles.heroSection}>
              <div className={styles.heroHeader}>
                <Icon iconName="Search" className={styles.heroIcon} />
                <div>
                  <h1 className={styles.heroTitle}>Search Policies</h1>
                  <p className={styles.heroSubtitle}>
                    Find policies, templates, packs, and compliance documents
                  </p>
                </div>
              </div>

              <div className={styles.searchBoxContainer}>
                <SearchBox
                  placeholder="Search by policy name, number, keywords, category..."
                  value={searchQuery}
                  onChange={(_, value) => this.setState({ searchQuery: value || '' })}
                  onSearch={this.handleSearch}
                  onClear={() => this.setState({ searchQuery: '', results: [], hasSearched: false })}
                  styles={{
                    root: { borderRadius: '6px', backgroundColor: '#ffffff' },
                    field: { fontSize: '16px', padding: '8px' },
                  }}
                />

                {/* Quick Filters */}
                <div className={styles.quickFilters}>
                  {typeOptions.map(opt => (
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

              {/* Recent Searches */}
              {!hasSearched && (
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

                  {/* Type Filter */}
                  <div className={styles.filterSection}>
                    <div className={styles.filterTitle}>Content Type</div>
                    {typeOptions.map(opt => (
                      <Checkbox
                        key={opt.key}
                        label={opt.text}
                        checked={filters.types.includes(opt.key)}
                        onChange={(_, checked) => this.handleFilterChange('types', opt.key, checked || false)}
                        className={styles.filterCheckbox}
                      />
                    ))}
                  </div>

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
                    {categoryOptions.slice(0, 6).map(opt => (
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
                      }))}
                      styles={{ root: { marginBottom: '8px' } }}
                    />
                    <DatePicker
                      placeholder="To date"
                      value={filters.dateTo || undefined}
                      onSelectDate={(date) => this.setState(prev => ({
                        filters: { ...prev.filters, dateTo: date || null }
                      }))}
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
                      <Spinner size={SpinnerSize.large} label="Searching..." />
                    </div>
                  ) : results.length > 0 ? (
                    <>
                      <div className={styles.resultsHeader}>
                        <span className={styles.resultsCount}>
                          {results.length} results for &ldquo;{searchQuery}&rdquo;
                        </span>
                        <Dropdown
                          placeholder="Sort by"
                          selectedKey={sortBy}
                          options={sortOptions}
                          onChange={(_, option) => this.setState({ sortBy: (option?.key as string) || 'relevance' })}
                          className={styles.sortDropdown}
                        />
                      </div>

                      {results.map(result => (
                        <div
                          key={`${result.type}-${result.id}`}
                          className={styles.resultCard}
                          onClick={() => this.handleResultClick(result)}
                        >
                          <div className={styles.resultHeader}>
                            <div
                              className={styles.resultIcon}
                              style={{ backgroundColor: getTypeColor(result.type) }}
                            >
                              <Icon iconName={getTypeIcon(result.type)} />
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

                          {result.highlights.length > 0 && (
                            <div className={styles.resultHighlights}>
                              {result.highlights[0]}
                            </div>
                          )}

                          <div className={styles.resultMeta}>
                            <span className={styles.resultMetaItem}>
                              <Icon iconName={getTypeIcon(result.type)} />
                              {result.type.charAt(0).toUpperCase() + result.type.slice(1)}
                            </span>
                            <span className={styles.resultMetaItem}>
                              <Icon iconName="Calendar" />
                              Modified {result.lastModified.toLocaleDateString()}
                            </span>
                            <span className={styles.resultMetaItem}>
                              <Icon iconName="CustomActivity" />
                              {Math.round(result.relevanceScore * 100)}% match
                            </span>
                          </div>
                        </div>
                      ))}
                    </>
                  ) : (
                    <div className={styles.noResults}>
                      <Icon iconName="SearchIssue" className={styles.noResultsIcon} />
                      <div className={styles.noResultsTitle}>No results found</div>
                      <div className={styles.noResultsText}>
                        We couldn&apos;t find anything matching &ldquo;{searchQuery}&rdquo;.
                        Try different keywords or adjust your filters.
                      </div>
                      <DefaultButton
                        text="Clear Filters"
                        onClick={this.clearFilters}
                        style={{ marginTop: '16px' }}
                      />
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
