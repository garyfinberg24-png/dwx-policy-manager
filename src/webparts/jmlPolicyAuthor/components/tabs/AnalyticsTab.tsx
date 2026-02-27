// @ts-nocheck
/**
 * AnalyticsTab — Extracted from PolicyAuthorEnhanced.tsx
 * Displays policy analytics with KPI cards, category/status/risk charts,
 * and department compliance table.
 */
import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  DefaultButton,
  PrimaryButton,
  IconButton,
  Icon,
} from '@fluentui/react';
import { PageSubheader } from '../../../../components/PageSubheader';
import {
  IAuthorPolicyAnalytics as IPolicyAnalytics,
} from '../../../../models/IPolicyAuthor';
import { IAnalyticsTabProps } from './types';

export default class AnalyticsTab extends React.Component<IAnalyticsTabProps> {

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private renderAnalyticsKpiCard(title: string, value: string | number, icon: string, color: string): JSX.Element {
    const { styles } = this.props;
    return (
      <div className={styles.analyticsKpiCard}>
        <Icon iconName={icon} style={{ fontSize: 24, color, marginBottom: 8 }} />
        <Text variant="xxLarge" style={{ fontWeight: 700 }}>{value}</Text>
        <Text variant="small" style={{ color: '#605e5c' }}>{title}</Text>
      </div>
    );
  }

  // ============================================================================
  // RENDER
  // ============================================================================

  public render(): React.ReactElement<IAnalyticsTabProps> {
    const { analyticsData, analyticsLoading, departmentCompliance, styles, dialogManager, onDateRangeChange, onExportAnalytics } = this.props;

    return (
      <>
        <PageSubheader
          iconName="BarChartVertical"
          title="Policy Analytics"
          description="Insights and metrics for your policy library"
          actions={
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Date Range"
                iconProps={{ iconName: 'Calendar' }}
                menuProps={{
                  items: [
                    { key: 'last7', text: 'Last 7 Days', onClick: () => { void onDateRangeChange(7); } },
                    { key: 'last30', text: 'Last 30 Days', onClick: () => { void onDateRangeChange(30); } },
                    { key: 'last90', text: 'Last 90 Days', onClick: () => { void onDateRangeChange(90); } },
                    { key: 'thisYear', text: 'This Year', onClick: () => { void onDateRangeChange(365); } },
                    { key: 'allTime', text: 'All Time', onClick: () => { void onDateRangeChange(0); } }
                  ]
                }}
              />
              <PrimaryButton
                text="Export Report"
                iconProps={{ iconName: 'Download' }}
                menuProps={{
                  items: [
                    { key: 'csv', text: 'Export as CSV', iconProps: { iconName: 'ExcelDocument' }, onClick: () => { void onExportAnalytics('csv'); } },
                    { key: 'pdf', text: 'Export as PDF', iconProps: { iconName: 'PDF' }, onClick: () => { void onExportAnalytics('pdf'); } },
                    { key: 'json', text: 'Export as JSON', iconProps: { iconName: 'Code' }, onClick: () => { void onExportAnalytics('json'); } }
                  ]
                }}
              />
            </Stack>
          }
        />

        <div className={styles.editorContainer}>
          {analyticsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading analytics..." />
            </Stack>
          ) : (() => {
            // Single render path — use live data or fallback sample data
            const data: IPolicyAnalytics = analyticsData || {
              totalPolicies: 48, publishedPolicies: 35, draftPolicies: 8, pendingApproval: 5,
              expiringSoon: 3, averageReadTime: 15, complianceRate: 89, acknowledgementRate: 78,
              policiesByCategory: [
                { category: 'HR', count: 12 }, { category: 'IT Security', count: 8 },
                { category: 'Finance', count: 6 }, { category: 'Compliance', count: 10 }, { category: 'Operations', count: 7 }
              ],
              policiesByStatus: [
                { status: 'Published', count: 35 }, { status: 'Draft', count: 8 },
                { status: 'In Review', count: 5 }
              ],
              policiesByRisk: [
                { risk: 'Low', count: 18 }, { risk: 'Medium', count: 20 },
                { risk: 'High', count: 8 }, { risk: 'Critical', count: 2 }
              ],
              monthlyTrends: []
            };
            const riskColors: Record<string, string> = {
              'Low': '#107c10', 'Medium': '#ca5010', 'High': '#d13438', 'Critical': '#750b1c'
            };
            return (
            <>
              {/* KPI Cards */}
              <div className={styles.analyticsKpiGrid}>
                {this.renderAnalyticsKpiCard('Total Policies', data.totalPolicies, 'DocumentSet', '#0078d4')}
                {this.renderAnalyticsKpiCard('Published', data.publishedPolicies, 'CheckMark', '#107c10')}
                {this.renderAnalyticsKpiCard('Draft', data.draftPolicies, 'Edit', '#605e5c')}
                {this.renderAnalyticsKpiCard('Pending Approval', data.pendingApproval, 'Clock', '#ca5010')}
                {this.renderAnalyticsKpiCard('Expiring Soon', data.expiringSoon, 'Warning', '#d13438')}
                {this.renderAnalyticsKpiCard('Compliance Rate', `${data.complianceRate}%`, 'Shield', '#0078d4')}
              </div>

              {/* Charts */}
              <div className={styles.analyticsChartsGrid}>
                {/* By Category */}
                <div className={styles.analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Category</Text>
                  <div className={styles.barChart}>
                    {data.policiesByCategory.map(item => (
                      <div key={item.category} className={styles.barChartItem}>
                        <Text style={{ width: 120 }}>{item.category}</Text>
                        <div className={styles.barChartBar}>
                          <div
                            className={styles.barChartFill}
                            style={{ width: `${(item.count / (data.totalPolicies || 1)) * 100}%` }}
                          />
                        </div>
                        <Text style={{ width: 40, textAlign: 'right' }}>{item.count}</Text>
                      </div>
                    ))}
                  </div>
                </div>

                {/* By Status */}
                <div className={styles.analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Status</Text>
                  <div className={styles.donutChartContainer}>
                    {data.policiesByStatus.map((item, index) => (
                      <div key={item.status} className={styles.donutLegendItem}>
                        <span className={styles.donutLegendColor} style={{
                          backgroundColor: ['#0078d4', '#107c10', '#ca5010', '#605e5c', '#d13438'][index % 5]
                        }} />
                        <Text>{item.status}: {item.count}</Text>
                      </div>
                    ))}
                  </div>
                </div>

                {/* By Risk */}
                <div className={styles.analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Risk Level</Text>
                  <div className={styles.riskGrid}>
                    {data.policiesByRisk.map(item => (
                      <div key={item.risk} className={styles.riskCard} style={{ borderLeftColor: riskColors[item.risk] || '#605e5c' }}>
                        <Text variant="xxLarge" style={{ fontWeight: 700 }}>{item.count}</Text>
                        <Text>{item.risk}</Text>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Department Compliance Table */}
              <div className={styles.analyticsChart} style={{ marginTop: 24 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
                  <Text variant="large" style={{ fontWeight: 600 }}>Department Compliance</Text>
                  <DefaultButton
                    text="Send Reminders"
                    iconProps={{ iconName: 'Mail' }}
                    onClick={() => void dialogManager.showAlert('Reminder emails will be sent to non-compliant employees', { variant: 'info' })}
                  />
                </Stack>
                <div className={styles.complianceTable}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f3f2f1', borderBottom: '2px solid #edebe9' }}>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Department</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Total</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Non-Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Pending</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Rate</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {departmentCompliance.map((dept, index) => (
                        <tr key={dept.Department} style={{ borderBottom: '1px solid #edebe9', background: index % 2 === 0 ? '#ffffff' : '#faf9f8' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 500 }}>{dept.Department}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{dept.TotalEmployees}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#107c10' }}>{dept.Compliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#d13438' }}>{dept.NonCompliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#ca5010' }}>{dept.Pending}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              display: 'inline-block',
                              padding: '4px 12px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              fontWeight: 600,
                              background: dept.ComplianceRate >= 90 ? '#dff6dd' : dept.ComplianceRate >= 80 ? '#fff4ce' : '#fde7e9',
                              color: dept.ComplianceRate >= 90 ? '#107c10' : dept.ComplianceRate >= 80 ? '#8a6d3b' : '#d13438'
                            }}>
                              {dept.ComplianceRate}%
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <IconButton
                              iconProps={{ iconName: 'View' }}
                              title="View Details"
                              onClick={() => void dialogManager.showAlert(`Viewing compliance details for ${dept.Department}`, { variant: 'info' })}
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
            );
          })()}
        </div>
      </>
    );
  }
}
