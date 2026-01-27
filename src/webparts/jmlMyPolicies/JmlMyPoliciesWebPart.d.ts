import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IJmlMyPoliciesWebPartProps {
    title: string;
    showComplianceScore: boolean;
    showPolicyPacks: boolean;
    showJMLIntegration: boolean;
}
export default class JmlMyPoliciesWebPart extends BaseClientSideWebPart<IJmlMyPoliciesWebPartProps> {
    private _isDarkTheme;
    private _sp;
    render(): void;
    protected onInit(): Promise<void>;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=JmlMyPoliciesWebPart.d.ts.map