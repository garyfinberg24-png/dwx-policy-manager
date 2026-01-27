import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

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
}
