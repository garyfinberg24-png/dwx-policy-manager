import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { DwxHubService } from '@dwx/core';

export interface IPolicyAuthorProps {
  title: string;
  enableTemplates: boolean;
  enableAutoSave: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
  /** DWx Hub service — undefined when Hub is unavailable (standalone mode) */
  dwxHub?: DwxHubService;
}
