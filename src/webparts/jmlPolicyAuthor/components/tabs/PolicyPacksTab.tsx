// @ts-nocheck
/**
 * PolicyPacksTab — Extracted from PolicyAuthorEnhanced.tsx
 * Displays policy pack management with stats summary and pack card grid.
 */
import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  PrimaryButton,
  IconButton,
  Icon,
} from '@fluentui/react';
import { PageSubheader } from '../../../../components/PageSubheader';
import { IPolicyPacksTabProps } from './types';

export default class PolicyPacksTab extends React.Component<IPolicyPacksTabProps> {

  public render(): React.ReactElement<IPolicyPacksTabProps> {
    const { policyPacks, policyPacksLoading, styles, dialogManager, onCreatePack } = this.props;

    return (
      <>
        <PageSubheader
          iconName="Package"
          title="Policy Packs"
          description="Manage bundled policy collections"
          actions={
            <PrimaryButton
              text="Create New Pack"
              iconProps={{ iconName: 'Add' }}
              onClick={onCreatePack}
            />
          }
        />

        <div className={styles.editorContainer}>
          {policyPacksLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy packs..." />
            </Stack>
          ) : policyPacks.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="Package" style={{ fontSize: 64, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="xLarge" style={{ fontWeight: 600 }}>No Policy Packs</Text>
              <Text style={{ color: '#605e5c', marginBottom: 24 }}>Create your first policy pack to bundle policies for easy distribution</Text>
              <PrimaryButton
                text="Create New Pack"
                iconProps={{ iconName: 'Add' }}
                onClick={onCreatePack}
              />
            </Stack>
          ) : (
            <>
              {/* Stats Summary */}
              <Stack horizontal tokens={{ childrenGap: 24 }} style={{ marginBottom: 24 }}>
                <div style={{ background: '#e8f4fd', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="Package" style={{ fontSize: 20, color: '#0078d4' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{policyPacks.length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Total Packs</Text>
                  </div>
                </div>
                <div style={{ background: '#dff6dd', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#107c10' }}>{policyPacks.filter(p => p.Status === 'Active').length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Active</Text>
                  </div>
                </div>
                <div style={{ background: '#fff4ce', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="Edit" style={{ fontSize: 20, color: '#8a6d3b' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#8a6d3b' }}>{policyPacks.filter(p => p.Status === 'Draft').length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Draft</Text>
                  </div>
                </div>
              </Stack>

              {/* Policy Pack Cards Grid */}
              <div className={styles.policyPackGrid}>
                {policyPacks.map(pack => (
                  <div key={pack.Id} className={styles.policyPackCard}>
                    <div className={styles.policyPackCardHeader}>
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                        <div>
                          <Text variant="large" style={{ fontWeight: 600, display: 'block' }}>{pack.Title}</Text>
                          <Text variant="small" style={{ color: '#605e5c', marginTop: 4, display: 'block' }}>{pack.Description}</Text>
                        </div>
                        <span style={{
                          display: 'inline-block',
                          padding: '4px 12px',
                          borderRadius: '12px',
                          fontSize: '11px',
                          fontWeight: 600,
                          textTransform: 'uppercase',
                          background: pack.Status === 'Active' ? '#dff6dd' : '#fff4ce',
                          color: pack.Status === 'Active' ? '#107c10' : '#8a6d3b'
                        }}>
                          {pack.Status}
                        </span>
                      </Stack>
                    </div>
                    <div className={styles.policyPackCardBody}>
                      <Stack horizontal tokens={{ childrenGap: 24 }}>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#0078d4', display: 'block' }}>{pack.PoliciesCount}</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Policies</Text>
                        </div>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#107c10', display: 'block' }}>{pack.AssignedTo}</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Assigned</Text>
                        </div>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#8a6d3b', display: 'block' }}>{pack.CompletionRate}%</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Complete</Text>
                        </div>
                      </Stack>
                      <div style={{ marginTop: 12 }}>
                        <div style={{ height: 6, background: '#f3f2f1', borderRadius: 3, overflow: 'hidden' }}>
                          <div style={{
                            height: '100%',
                            width: `${pack.CompletionRate}%`,
                            background: pack.CompletionRate >= 80 ? '#107c10' : pack.CompletionRate >= 50 ? '#ca5010' : '#d13438',
                            borderRadius: 3,
                            transition: 'width 0.3s ease'
                          }} />
                        </div>
                      </div>
                    </div>
                    <div className={styles.policyPackCardFooter}>
                      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                        <Icon iconName="People" style={{ fontSize: 14, color: '#605e5c' }} />
                        <Text variant="small" style={{ color: '#605e5c' }}>{pack.TargetAudience}</Text>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <IconButton
                          iconProps={{ iconName: 'Edit' }}
                          title="Edit Pack"
                          onClick={() => void dialogManager.showAlert(`Edit pack: ${pack.Title}`, { variant: 'info' })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'View' }}
                          title="View Details"
                          onClick={() => void dialogManager.showAlert(`View details for: ${pack.Title}`, { variant: 'info' })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Send' }}
                          title="Assign Pack"
                          onClick={() => void dialogManager.showAlert(`Assign pack: ${pack.Title}`, { variant: 'info' })}
                        />
                      </Stack>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </>
    );
  }
}
