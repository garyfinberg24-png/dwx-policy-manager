import * as React from 'react';
import { IMyPoliciesProps } from './IMyPoliciesProps';
import { IPersonalPolicyView } from '../../../models/IPolicy';
export interface IMyPoliciesState {
    loading: boolean;
    error: string | null;
    personalView: IPersonalPolicyView | null;
    refreshing: boolean;
}
export default class MyPolicies extends React.Component<IMyPoliciesProps, IMyPoliciesState> {
    private policyPackService;
    constructor(props: IMyPoliciesProps);
    componentDidMount(): Promise<void>;
    private loadPersonalView;
    private handleRefresh;
    private handleAcknowledge;
    private renderComplianceScore;
    private renderUrgentPolicies;
    private renderDueSoonPolicies;
    private renderPolicyPacks;
    private renderJMLIntegration;
    private renderEmptyState;
    render(): React.ReactElement<IMyPoliciesProps>;
}
//# sourceMappingURL=MyPolicies.d.ts.map