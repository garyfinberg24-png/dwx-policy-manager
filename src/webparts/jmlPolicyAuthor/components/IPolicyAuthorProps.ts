import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

export interface IPolicyAuthorProps {
  title: string;
  enableTemplates: boolean;
  enableAutoSave: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
}
