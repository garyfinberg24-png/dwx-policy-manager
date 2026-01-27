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

import * as strings from 'JmlPolicyAdminWebPartStrings';
import PolicyAdmin from './components/PolicyAdmin';
import { IPolicyAdminProps } from './components/IPolicyAdminProps';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';

export interface IDwxPolicyAdminWebPartProps {
  title: string;
  showAuditLog: boolean;
  enableBulkOperations: boolean;
}

export default class DwxPolicyAdminWebPart extends BaseClientSideWebPart<IDwxPolicyAdminWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IPolicyAdminProps> = React.createElement(
      PolicyAdmin,
      {
        title: this.properties.title,
        showAuditLog: this.properties.showAuditLog,
        enableBulkOperations: this.properties.enableBulkOperations,
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
                PropertyPaneToggle('showAuditLog', {
                  label: 'Show Audit Log',
                  checked: true
                }),
                PropertyPaneToggle('enableBulkOperations', {
                  label: 'Enable Bulk Operations',
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
