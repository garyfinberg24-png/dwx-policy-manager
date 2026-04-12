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
import { injectSharePointOverrides, signalAppReady } from '../../utils/SharePointOverrides';
import { ThemeManager } from '../../utils/themeManager';

export interface IDwxStartScreenWebPartProps {
  title: string;
}

export default class DwxStartScreenWebPart extends BaseClientSideWebPart<IDwxStartScreenWebPartProps> {
  private _sp: SPFI;

  public render(): void {
    // Detect role from localStorage (set by JmlAppLayout / PolicyManagerHeader)
    const storedRole = localStorage.getItem('pm_detected_role') || 'User';

    // Render directly into domElement with inline fallback
    this.domElement.innerHTML = '';
    this.domElement.style.cssText = 'min-height: 100vh; width: 100%;';

    const container = document.createElement('div');
    container.id = 'dwx-start-screen-root';
    container.style.cssText = 'min-height: 100vh; width: 100%;';
    this.domElement.appendChild(container);

    const element = React.createElement(
      StartScreen,
      {
        sp: this._sp,
        userName: this.context.pageContext.user.displayName || 'User',
        userRole: storedRole,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        onDismiss: () => {
          sessionStorage.setItem('pm_start_dismissed', 'true');
          window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/PolicyHub.aspx`;
        }
      }
    );

    ReactDom.render(element, container);

    // Signal app ready — hides the JML loading skeleton
    signalAppReady();
    console.log('[DwxStartScreen] Rendered and signalled ready');
  }

  protected async onInit(): Promise<void> {
    // Inject critical CSS immediately to hide SP chrome before React mounts
    injectSharePointOverrides();

    // Also inject inline style to hide SP skeleton ASAP
    const style = document.createElement('style');
    style.textContent = `
      .CanvasZone [class*="placeholder"],
      .CanvasZone [class*="Placeholder"],
      [class*="webPartLoading"],
      [class*="sp-webpart-loading"] {
        display: none !important;
      }
      .CanvasZone { overflow: visible !important; }
    `;
    document.head.appendChild(style);

    await super.onInit();
    this._sp = getSP(this.context);

    // Load and apply the saved theme (same as JmlAppLayout does)
    try {
      const stored = ThemeManager.getTheme();
      if (stored && stored.primaryColor) {
        ThemeManager.apply(stored);
      } else {
        // Try loading from SP
        const spTheme = await ThemeManager.loadFromSP(this._sp);
        if (spTheme && spTheme.primaryColor) {
          ThemeManager.apply(spTheme);
        }
      }
    } catch { /* use defaults if theme load fails */ }
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
