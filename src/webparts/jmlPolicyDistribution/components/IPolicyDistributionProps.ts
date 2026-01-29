import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

export interface IPolicyDistributionProps {
  title: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
}
