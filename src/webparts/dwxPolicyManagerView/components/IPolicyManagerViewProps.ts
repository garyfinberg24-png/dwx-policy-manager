import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

export interface IPolicyManagerViewProps {
  title: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
}
