import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
export interface IMyPoliciesProps {
    title: string;
    showComplianceScore: boolean;
    showPolicyPacks: boolean;
    showJMLIntegration: boolean;
    isDarkTheme: boolean;
    hasTeamsContext: boolean;
    sp: SPFI;
    context: WebPartContext;
}
//# sourceMappingURL=IMyPoliciesProps.d.ts.map