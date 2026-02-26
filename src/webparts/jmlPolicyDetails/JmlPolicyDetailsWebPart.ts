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

import * as strings from 'JmlPolicyDetailsWebPartStrings';
import PolicyDetails from './components/PolicyDetails';
import { IPolicyDetailsProps } from './components/IPolicyDetailsProps';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';
import { DwxHubService } from '@dwx/core';

export interface IDwxPolicyDetailsWebPartProps {
  title: string;
  showRelatedDocuments: boolean;
  showComments: boolean;
  showRatings: boolean;
  enableQuiz: boolean;
}

export default class DwxPolicyDetailsWebPart extends BaseClientSideWebPart<IDwxPolicyDetailsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _sp: SPFI;
  private _dwxHub: DwxHubService | undefined;

  public render(): void {
    const element: React.ReactElement<IPolicyDetailsProps> = React.createElement(
      PolicyDetails,
      {
        title: this.properties.title,
        showRelatedDocuments: this.properties.showRelatedDocuments,
        showComments: this.properties.showComments,
        showRatings: this.properties.showRatings,
        enableQuiz: this.properties.enableQuiz,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        sp: this._sp,
        context: this.context,
        dwxHub: this._dwxHub
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    injectSharePointOverrides();
    await super.onInit();
    this._sp = getSP(this.context);

    // DWx Core integration (graceful degradation — app works without Hub)
    try {
      this._dwxHub = new DwxHubService(this.context, this._sp);
      if (!await this._dwxHub.isHubAvailable()) {
        this._dwxHub = undefined;
      }
    } catch (err) {
      console.warn('[PolicyDetails] DWx Hub unavailable, running standalone:', err);
      this._dwxHub = undefined;
    }
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
                PropertyPaneToggle('showRelatedDocuments', {
                  label: 'Show Related Documents',
                  checked: true
                }),
                PropertyPaneToggle('showComments', {
                  label: 'Show Comments',
                  checked: true
                }),
                PropertyPaneToggle('showRatings', {
                  label: 'Show Ratings',
                  checked: true
                }),
                PropertyPaneToggle('enableQuiz', {
                  label: 'Enable Quiz',
                  checked: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
