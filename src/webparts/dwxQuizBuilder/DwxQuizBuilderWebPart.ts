import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DwxQuizBuilderWebPartStrings';
import { QuizBuilderWrapper } from './components/QuizBuilderWrapper';
import { IQuizBuilderWrapperProps } from './components/IQuizBuilderWrapperProps';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';

export interface IDwxQuizBuilderWebPartProps {
  title: string;
  enableQuestionBanks: boolean;
  enableImportExport: boolean;
  aiFunctionUrl: string;
}

export default class DwxQuizBuilderWebPart extends BaseClientSideWebPart<IDwxQuizBuilderWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IQuizBuilderWrapperProps> = React.createElement(
      QuizBuilderWrapper,
      {
        title: this.properties.title,
        enableQuestionBanks: this.properties.enableQuestionBanks,
        enableImportExport: this.properties.enableImportExport,
        aiFunctionUrl: this.properties.aiFunctionUrl || '',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        sp: this._sp,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    injectSharePointOverrides();
    await super.onInit();
    this._sp = getSP(this.context);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneToggle('enableQuestionBanks', {
                  label: 'Enable Question Banks',
                  checked: true
                }),
                PropertyPaneToggle('enableImportExport', {
                  label: 'Enable Import/Export',
                  checked: true
                }),
                PropertyPaneTextField('aiFunctionUrl', {
                  label: 'AI Function URL',
                  description: 'Full URL to the Azure Function for AI quiz generation (including ?code= key)',
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
