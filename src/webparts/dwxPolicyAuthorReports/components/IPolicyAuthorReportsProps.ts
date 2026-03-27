import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

export interface IPolicyAuthorReportsProps {
  title: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
}
