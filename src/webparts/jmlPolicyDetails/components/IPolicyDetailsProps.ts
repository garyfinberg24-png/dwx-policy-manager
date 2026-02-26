import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { DwxHubService } from '@dwx/core';

export interface IPolicyDetailsProps {
  title: string;
  showRelatedDocuments: boolean;
  showComments: boolean;
  showRatings: boolean;
  enableQuiz: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
  dwxHub?: DwxHubService;
}
