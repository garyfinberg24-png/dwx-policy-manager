// @ts-nocheck
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DwxStartScreenWebPartStrings';
import { StartScreen } from '../../components/StartScreen/StartScreen';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../utils/pnpConfig';
import { injectSharePointOverrides } from '../../utils/SharePointOverrides';

export interface IDwxStartScreenWebPartProps {
  title: string;
}

export default class DwxStartScreenWebPart extends BaseClientSideWebPart<IDwxStartScreenWebPartProps> {
  private _sp: SPFI;

  public render(): void {
    // Detect role from localStorage (set by JmlAppLayout / PolicyManagerHeader)
    const storedRole = localStorage.getItem('pm_detected_role') || 'User';

    const element = React.createElement(
      StartScreen,
      {
        sp: this._sp,
        userName: this.context.pageContext.user.displayName || 'User',
        userRole: storedRole,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        onDismiss: () => {
          // Navigate to Policy Hub when user clicks "Skip to Policy Hub"
          sessionStorage.setItem('pm_start_dismissed', 'true');
          window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/PolicyHub.aspx`;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Hide SharePoint chrome for full app-like experience
    injectSharePointOverrides();
    await super.onInit();
    this._sp = getSP(this.context);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
