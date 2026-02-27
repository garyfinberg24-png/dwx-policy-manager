// @ts-nocheck
/**
 * DelegationsTab — Extracted from PolicyAuthorEnhanced.tsx
 * Displays policy delegation requests with KPI cards and delegation list.
 */
import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  DefaultButton,
  PrimaryButton,
  Icon,
} from '@fluentui/react';
import { PageSubheader } from '../../../../components/PageSubheader';
import { IDelegationsTabProps } from './types';

export default class DelegationsTab extends React.Component<IDelegationsTabProps> {

  public render(): React.ReactElement<IDelegationsTabProps> {
    const { delegatedRequests, delegationsLoading, delegationKpis, styles, onNewDelegation, onStartPolicy } = this.props;

    return (
      <>
        <PageSubheader
          iconName="Assign"
          title="Policy Delegations"
          description="Policies delegated to you for creation"
          actions={
            <PrimaryButton
              text="New Delegation"
              iconProps={{ iconName: 'Add' }}
              onClick={onNewDelegation}
            />
          }
        />

        {/* KPI Summary Cards */}
        <div className={styles.delegationKpiGrid}>
          <div className={styles.delegationKpiCard}>
            <div className={styles.delegationKpiIcon} style={{ background: '#e8f4fd' }}>
              <Icon iconName="Assign" style={{ fontSize: 20, color: '#0078d4' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{delegationKpis.activeDelegations}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Active Delegations</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard}>
            <div className={styles.delegationKpiIcon} style={{ background: '#dff6dd' }}>
              <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#107c10' }}>{delegationKpis.completedThisMonth}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Completed This Month</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard}>
            <div className={styles.delegationKpiIcon} style={{ background: '#fff4ce' }}>
              <Icon iconName="Clock" style={{ fontSize: 20, color: '#8a6d3b' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#8a6d3b' }}>{delegationKpis.averageCompletionTime}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Avg. Completion Time</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard}>
            <div className={styles.delegationKpiIcon} style={{ background: '#fde7e9' }}>
              <Icon iconName="Warning" style={{ fontSize: 20, color: '#d13438' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#d13438' }}>{delegationKpis.overdue}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Overdue</Text>
            </div>
          </div>
        </div>

        <div className={styles.editorContainer}>
          {delegationsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading delegations..." />
            </Stack>
          ) : delegatedRequests.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="Assign" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No delegated policies</Text>
              <Text>You don't have any policy creation requests assigned to you</Text>
            </Stack>
          ) : (
            <div className={styles.delegationList}>
              {delegatedRequests.map(request => (
                <div key={request.Id} className={styles.delegationCard}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <div>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{request.Title}</Text>
                      <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
                        Requested by {request.RequestedBy} • {request.PolicyType}
                      </Text>
                      <Text variant="small" style={{ marginTop: 8 }}>{request.Description}</Text>
                    </div>
                    <Stack horizontalAlign="end">
                      <span className={styles.urgencyBadge} data-urgency={request.Urgency}>
                        {request.Urgency}
                      </span>
                      <Text variant="small" style={{ color: '#605e5c', marginTop: 8 }}>
                        Due: {new Date(request.DueDate).toLocaleDateString()}
                      </Text>
                    </Stack>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
                    <PrimaryButton
                      text="Start Policy"
                      iconProps={{ iconName: 'Add' }}
                      onClick={() => onStartPolicy(request)}
                    />
                    <DefaultButton
                      text="View Details"
                      iconProps={{ iconName: 'Info' }}
                    />
                  </Stack>
                </div>
              ))}
            </div>
          )}
        </div>
      </>
    );
  }
}
