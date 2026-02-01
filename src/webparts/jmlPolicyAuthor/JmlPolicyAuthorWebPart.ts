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

import * as strings from 'JmlPolicyAuthorWebPartStrings';
import PolicyAuthorEnhanced from './components/PolicyAuthorEnhanced';
import { IPolicyAuthorProps } from './components/IPolicyAuthorProps';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';
import { DwxHubService, DwxAppRegistryService } from '@dwx/core';

export interface IDwxPolicyAuthorWebPartProps {
  title: string;
  enableTemplates: boolean;
  enableAutoSave: boolean;
}

export default class DwxPolicyAuthorWebPart extends BaseClientSideWebPart<IDwxPolicyAuthorWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _sp: SPFI;
  private _dwxHub: DwxHubService | undefined;

  public render(): void {
    const element: React.ReactElement<IPolicyAuthorProps> = React.createElement(
      PolicyAuthorEnhanced,
      {
        title: this.properties.title,
        enableTemplates: this.properties.enableTemplates,
        enableAutoSave: this.properties.enableAutoSave,
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

    // DWx Core integration (graceful degradation)
    try {
      this._dwxHub = new DwxHubService(this.context, this._sp);
      if (await this._dwxHub.isHubAvailable()) {
        const registry = new DwxAppRegistryService(this._dwxHub);
        await registry.heartbeat('PolicyManager', '1.2.1');
      }
    } catch (err) {
      console.warn('[PolicyAuthor] DWx Hub unavailable, running standalone:', err);
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
                PropertyPaneToggle('enableTemplates', {
                  label: 'Enable Templates',
                  checked: true
                }),
                PropertyPaneToggle('enableAutoSave', {
                  label: 'Enable Auto-Save',
                  checked: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
