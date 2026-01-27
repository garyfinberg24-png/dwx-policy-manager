import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IQuizBuilderWrapperProps {
  title: string;
  enableQuestionBanks: boolean;
  enableImportExport: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  sp: SPFI;
  context: WebPartContext;
}
