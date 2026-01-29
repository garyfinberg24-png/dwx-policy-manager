import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'JmlPolicyHubWebPartStrings';
import PolicyHub from './components/PolicyHub';
import { IPolicyHubProps } from './components/IPolicyHubProps';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';

export interface IDwxPolicyHubWebPartProps {
  title: string;
  showDocumentCenter: boolean;
  enableAdvancedSearch: boolean;
  itemsPerPage: number;
  showFacets: boolean;
  enableFeaturedPolicies: boolean;
  enableRecentlyViewed: boolean;
}

export default class DwxPolicyHubWebPart extends BaseClientSideWebPart<IDwxPolicyHubWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IPolicyHubProps> = React.createElement(
      PolicyHub,
      {
        title: this.properties.title,
        showDocumentCenter: this.properties.showDocumentCenter,
        enableAdvancedSearch: this.properties.enableAdvancedSearch,
        itemsPerPage: this.properties.itemsPerPage,
        showFacets: this.properties.showFacets,
        enableFeaturedPolicies: this.properties.enableFeaturedPolicies === true,
        enableRecentlyViewed: this.properties.enableRecentlyViewed === true,
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
                PropertyPaneToggle('showDocumentCenter', {
                  label: 'Show Document Center',
                  checked: true
                }),
                PropertyPaneToggle('enableAdvancedSearch', {
                  label: 'Enable Advanced Search',
                  checked: true
                }),
                PropertyPaneToggle('showFacets', {
                  label: 'Show Faceted Filters',
                  checked: true
                }),
                PropertyPaneToggle('enableFeaturedPolicies', {
                  label: 'Show Featured Policies Section',
                  checked: true
                }),
                PropertyPaneToggle('enableRecentlyViewed', {
                  label: 'Show Recently Viewed Section',
                  checked: true
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: 'Items Per Page',
                  min: 10,
                  max: 100,
                  step: 10,
                  value: 20
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
