import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { DwxHubService } from '@dwx/core';

export interface IPolicyHubProps {
  title: string;
  showDocumentCenter: boolean;
  enableAdvancedSearch: boolean;
  itemsPerPage: number;
  showFacets: boolean;
  enableFeaturedPolicies: boolean;
  enableRecentlyViewed: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
  /** DWx Hub service — undefined when Hub is unavailable (standalone mode) */
  dwxHub?: DwxHubService;
}
