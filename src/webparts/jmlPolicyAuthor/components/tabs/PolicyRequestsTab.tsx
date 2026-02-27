// @ts-nocheck
/**
 * PolicyRequestsTab — Extracted from PolicyAuthorEnhanced.tsx
 * Displays policy creation requests submitted by managers, with KPI cards,
 * status filter chips, request cards, and a detail fly-in panel.
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
  Panel,
  PanelType,
} from '@fluentui/react';
import { PageSubheader } from '../../../../components/PageSubheader';
import {
  IPolicyAuthorRequest as IPolicyRequest,
} from '../../../../models/IPolicyAuthor';
import { IPolicyRequestsTabProps } from './types';

export default class PolicyRequestsTab extends React.Component<IPolicyRequestsTabProps> {

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private getRequestStatusColor(status: string): string {
    switch (status) {
      case 'New': return '#0078d4';
      case 'Assigned': return '#8764b8';
      case 'InProgress': return '#f59e0b';
      case 'Draft Ready': return '#14b8a6';
      case 'Completed': return '#107c10';
      case 'Rejected': return '#d13438';
      default: return '#605e5c';
    }
  }

  private getPriorityColor(priority: string): string {
    switch (priority) {
      case 'Critical': return '#d13438';
      case 'High': return '#f97316';
      case 'Medium': return '#f59e0b';
      case 'Low': return '#64748b';
      default: return '#605e5c';
    }
  }

  private getSamplePolicyRequests(): IPolicyRequest[] {
    return [
      {
        Id: 1, Title: 'Data Retention Policy for Cloud Storage',
        RequestedBy: 'Sarah Mitchell', RequestedByEmail: 'sarah.mitchell@company.com', RequestedByDepartment: 'IT Security',
        PolicyCategory: 'IT Security', PolicyType: 'New Policy', Priority: 'High',
        TargetAudience: 'All IT Staff, Development Teams', BusinessJustification: 'New GDPR requirements mandate clear data retention guidelines for all cloud storage services including Azure Blob, AWS S3, and Google Cloud Storage. Without this policy we are at risk of non-compliance.',
        RegulatoryDriver: 'GDPR Article 5(1)(e) — Storage Limitation', DesiredEffectiveDate: '2026-03-01', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'Please reference the existing Data Classification Policy and align retention periods accordingly. Legal has reviewed the requirements.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-27T09:15:00Z', Modified: '2026-01-27T09:15:00Z'
      },
      {
        Id: 2, Title: 'Remote Work Equipment & Ergonomics Policy',
        RequestedBy: 'James Thornton', RequestedByEmail: 'james.thornton@company.com', RequestedByDepartment: 'Human Resources',
        PolicyCategory: 'HR Policies', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'All Remote & Hybrid Employees', BusinessJustification: 'With 60% of workforce now remote, we need formal guidelines on equipment provisioning, ergonomic assessments, and home office stipend eligibility.',
        RegulatoryDriver: 'Health & Safety at Work Act', DesiredEffectiveDate: '2026-04-01', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Facilities team can provide ergonomic assessment checklist template.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-26T14:30:00Z', Modified: '2026-01-26T14:30:00Z'
      },
      {
        Id: 3, Title: 'AI & Machine Learning Usage Policy',
        RequestedBy: 'Dr. Aisha Patel', RequestedByEmail: 'aisha.patel@company.com', RequestedByDepartment: 'Innovation',
        PolicyCategory: 'IT Security', PolicyType: 'New Policy', Priority: 'Critical',
        TargetAudience: 'All Employees', BusinessJustification: 'Employees are using ChatGPT, Copilot, and other AI tools without guidelines. We need clear policy on acceptable use, data handling, intellectual property, and prohibited use cases.',
        RegulatoryDriver: 'EU AI Act, Internal IP Protection', DesiredEffectiveDate: '2026-02-15', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'Legal and InfoSec have drafted initial talking points. Board has flagged this as urgent. Please prioritise.',
        AttachmentUrls: [], Status: 'Assigned', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-20T11:00:00Z', Modified: '2026-01-22T08:45:00Z'
      },
      {
        Id: 4, Title: 'Vendor Risk Assessment Policy Update',
        RequestedBy: 'Robert Kumar', RequestedByEmail: 'robert.kumar@company.com', RequestedByDepartment: 'Procurement',
        PolicyCategory: 'Compliance', PolicyType: 'Policy Update', Priority: 'High',
        TargetAudience: 'Procurement, Legal, IT Security', BusinessJustification: 'Current vendor assessment policy is 2 years old and does not cover SaaS vendor risks, supply chain security, or ESG requirements.',
        RegulatoryDriver: 'ISO 27001, SOC 2 Type II', DesiredEffectiveDate: '2026-03-15', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Attach current vendor assessment checklist for reference. Procurement team available for consultation.',
        AttachmentUrls: [], Status: 'InProgress', AssignedAuthor: 'Mark Davies', AssignedAuthorEmail: 'mark.davies@company.com', Created: '2026-01-15T10:00:00Z', Modified: '2026-01-25T16:30:00Z'
      },
      {
        Id: 5, Title: 'Employee Social Media Conduct Policy',
        RequestedBy: 'Emma Whitfield', RequestedByEmail: 'emma.whitfield@company.com', RequestedByDepartment: 'Marketing',
        PolicyCategory: 'HR Policies', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'All Employees', BusinessJustification: 'Recent incidents of employees posting confidential project information on LinkedIn. Need clear guidelines on what can and cannot be shared on social media regarding company business.',
        RegulatoryDriver: 'Confidentiality & NDA Compliance', DesiredEffectiveDate: '2026-04-15', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Marketing has a brand guidelines document that should be referenced.',
        AttachmentUrls: [], Status: 'Completed', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-05T09:00:00Z', Modified: '2026-01-24T14:15:00Z'
      },
      {
        Id: 6, Title: 'Incident Response & Breach Notification Policy',
        RequestedBy: 'Sarah Mitchell', RequestedByEmail: 'sarah.mitchell@company.com', RequestedByDepartment: 'IT Security',
        PolicyCategory: 'IT Security', PolicyType: 'Policy Update', Priority: 'Critical',
        TargetAudience: 'IT Security, Management, Legal', BusinessJustification: 'Our incident response policy was written pre-cloud migration. Need to update for hybrid infrastructure, include cloud-specific playbooks, and align with 72-hour GDPR breach notification window.',
        RegulatoryDriver: 'GDPR Article 33 & 34, NIS2 Directive', DesiredEffectiveDate: '2026-02-28', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'CISO wants this prioritised. Include tabletop exercise requirements.',
        AttachmentUrls: [], Status: 'Assigned', AssignedAuthor: 'Mark Davies', AssignedAuthorEmail: 'mark.davies@company.com', Created: '2026-01-18T08:30:00Z', Modified: '2026-01-21T11:00:00Z'
      },
      {
        Id: 7, Title: 'Parental Leave & Return-to-Work Policy',
        RequestedBy: 'James Thornton', RequestedByEmail: 'james.thornton@company.com', RequestedByDepartment: 'Human Resources',
        PolicyCategory: 'HR Policies', PolicyType: 'Policy Update', Priority: 'Low',
        TargetAudience: 'All Employees', BusinessJustification: 'UK government has updated shared parental leave entitlements. Our policy needs to reflect new statutory minimums and company-enhanced provisions.',
        RegulatoryDriver: 'Employment Rights Act 1996 (updated)', DesiredEffectiveDate: '2026-06-01', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'HR Legal counsel has reviewed the statutory changes. Draft available.',
        AttachmentUrls: [], Status: 'Draft Ready', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-10T13:00:00Z', Modified: '2026-01-28T10:00:00Z'
      },
      {
        Id: 8, Title: 'Environmental Sustainability & Carbon Reporting Policy',
        RequestedBy: 'Olivia Green', RequestedByEmail: 'olivia.green@company.com', RequestedByDepartment: 'Operations',
        PolicyCategory: 'Environmental', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'Operations, Facilities, Finance', BusinessJustification: 'New CSRD (Corporate Sustainability Reporting Directive) requirements mean we need a formal sustainability policy covering carbon reporting, waste management, and supply chain environmental standards.',
        RegulatoryDriver: 'CSRD, TCFD, UK Energy Savings Opportunity Scheme', DesiredEffectiveDate: '2026-05-01', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'ESG consultants have provided a framework document. Finance team needs to be involved for carbon accounting.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-28T16:00:00Z', Modified: '2026-01-28T16:00:00Z'
      }
    ];
  }

  // ============================================================================
  // RENDER
  // ============================================================================

  public render(): React.ReactElement<IPolicyRequestsTabProps> {
    const { policyRequests, policyRequestsLoading, requestStatusFilter, selectedPolicyRequest, showPolicyRequestDetailPanel, styles, context, onSetState, onCreatePolicyFromRequest } = this.props;

    const statusFilters = ['All', 'New', 'Assigned', 'InProgress', 'Draft Ready', 'Completed', 'Rejected'];
    const filteredRequests = requestStatusFilter === 'All' ? policyRequests : policyRequests.filter(r => r.Status === requestStatusFilter);

    // KPI counts
    const newCount = policyRequests.filter(r => r.Status === 'New').length;
    const assignedCount = policyRequests.filter(r => r.Status === 'Assigned').length;
    const inProgressCount = policyRequests.filter(r => r.Status === 'InProgress').length;
    const completedCount = policyRequests.filter(r => r.Status === 'Completed' || r.Status === 'Draft Ready').length;
    const criticalCount = policyRequests.filter(r => r.Priority === 'Critical' && r.Status !== 'Completed').length;

    return (
      <>
        <PageSubheader
          iconName="PageAdd"
          title="Policy Requests"
          description="Review and manage policy creation requests submitted by managers"
          actions={
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={() => onSetState({ policyRequests: this.getSamplePolicyRequests() })} />
            </Stack>
          }
        />

        {/* KPI Summary Cards — including Critical as a card */}
        <div className={styles.delegationKpiGrid}>
          <div className={styles.delegationKpiCard} onClick={() => onSetState({ requestStatusFilter: 'New' })} style={{ cursor: 'pointer' }}>
            <div className={styles.delegationKpiIcon} style={{ background: '#e8f4fd' }}>
              <Icon iconName="NewMail" style={{ fontSize: 20, color: '#0078d4' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{newCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>New Requests</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard} onClick={() => onSetState({ requestStatusFilter: 'Assigned' })} style={{ cursor: 'pointer' }}>
            <div className={styles.delegationKpiIcon} style={{ background: '#f3eefc' }}>
              <Icon iconName="People" style={{ fontSize: 20, color: '#8764b8' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#8764b8' }}>{assignedCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Assigned</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard} onClick={() => onSetState({ requestStatusFilter: 'InProgress' })} style={{ cursor: 'pointer' }}>
            <div className={styles.delegationKpiIcon} style={{ background: '#fff8e6' }}>
              <Icon iconName="Edit" style={{ fontSize: 20, color: '#f59e0b' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#f59e0b' }}>{inProgressCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>In Progress</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard} onClick={() => onSetState({ requestStatusFilter: 'All' })} style={{ cursor: 'pointer' }}>
            <div className={styles.delegationKpiIcon} style={{ background: '#dff6dd' }}>
              <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#107c10' }}>{completedCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Completed</Text>
            </div>
          </div>
          <div className={styles.delegationKpiCard} onClick={() => onSetState({ requestStatusFilter: 'All' })} style={{ cursor: 'pointer' }}>
            <div className={styles.delegationKpiIcon} style={{ background: '#fef2f2' }}>
              <Icon iconName="ShieldAlert" style={{ fontSize: 20, color: '#d13438' }} />
            </div>
            <div className={styles.delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#d13438' }}>{criticalCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Critical</Text>
            </div>
          </div>
        </div>

        {/* Status Filter Chips */}
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {statusFilters.map(status => (
            <DefaultButton
              key={status}
              text={status === 'All' ? `All (${policyRequests.length})` : `${status} (${policyRequests.filter(r => r.Status === status).length})`}
              checked={requestStatusFilter === status}
              styles={{
                root: {
                  borderRadius: 20,
                  minWidth: 'auto',
                  padding: '2px 14px',
                  height: 32,
                  border: requestStatusFilter === status ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: requestStatusFilter === status ? '#f0fdfa' : 'transparent',
                  color: requestStatusFilter === status ? '#0d9488' : '#605e5c',
                  fontWeight: requestStatusFilter === status ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => onSetState({ requestStatusFilter: status })}
            />
          ))}
        </Stack>

        <div className={styles.editorContainer}>
          {policyRequestsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy requests..." />
            </Stack>
          ) : filteredRequests.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="PageAdd" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No policy requests</Text>
              <Text>No requests match the selected filter</Text>
            </Stack>
          ) : (
            <div className={styles.delegationList}>
              {filteredRequests.map(request => (
                <div
                  key={request.Id}
                  className={styles.delegationCard}
                  style={{ cursor: 'pointer', borderLeft: `4px solid ${this.getPriorityColor(request.Priority)}` }}
                  onClick={() => onSetState({ selectedPolicyRequest: request, showPolicyRequestDetailPanel: true })}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <div style={{ flex: 1 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{request.Title}</Text>
                        {request.Priority === 'Critical' && (
                          <span style={{ background: '#fde7e9', color: '#d13438', padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 700, textTransform: 'uppercase' as const }}>CRITICAL</span>
                        )}
                      </Stack>
                      <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
                        Requested by <strong>{request.RequestedBy}</strong> ({request.RequestedByDepartment}) &bull; {request.PolicyCategory} &bull; {request.PolicyType}
                      </Text>
                      <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130' }}>
                        {request.BusinessJustification.length > 150 ? request.BusinessJustification.substring(0, 150) + '...' : request.BusinessJustification}
                      </Text>
                      <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
                        <Text variant="small" style={{ color: '#605e5c' }}>
                          <Icon iconName="Calendar" style={{ marginRight: 4, fontSize: 12 }} />
                          Target: {new Date(request.DesiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}
                        </Text>
                        <Text variant="small" style={{ color: '#605e5c' }}>
                          <Icon iconName="Clock" style={{ marginRight: 4, fontSize: 12 }} />
                          Read within: {request.ReadTimeframeDays} days
                        </Text>
                        {request.RequiresAcknowledgement && (
                          <Text variant="small" style={{ color: '#0d9488' }}>
                            <Icon iconName="CheckboxComposite" style={{ marginRight: 4, fontSize: 12 }} /> Acknowledgement
                          </Text>
                        )}
                        {request.RequiresQuiz && (
                          <Text variant="small" style={{ color: '#8764b8' }}>
                            <Icon iconName="Questionnaire" style={{ marginRight: 4, fontSize: 12 }} /> Quiz Required
                          </Text>
                        )}
                      </Stack>
                    </div>
                    <Stack horizontalAlign="end" tokens={{ childrenGap: 4 }}>
                      <span style={{
                        background: `${this.getRequestStatusColor(request.Status)}15`,
                        color: this.getRequestStatusColor(request.Status),
                        padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
                      }}>
                        {request.Status === 'InProgress' ? 'In Progress' : request.Status}
                      </span>
                      <Text variant="tiny" style={{ color: '#a19f9d', marginTop: 4 }}>
                        {new Date(request.Created).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                      </Text>
                      {request.AssignedAuthor && (
                        <Text variant="tiny" style={{ color: '#605e5c' }}>
                          <Icon iconName="Contact" style={{ marginRight: 2, fontSize: 10 }} /> {request.AssignedAuthor}
                        </Text>
                      )}
                    </Stack>
                  </Stack>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Policy Request Detail Panel */}
        {showPolicyRequestDetailPanel && selectedPolicyRequest && (
          <Panel
            isOpen={showPolicyRequestDetailPanel}
            onDismiss={() => onSetState({ showPolicyRequestDetailPanel: false, selectedPolicyRequest: null })}
            type={PanelType.medium}
            headerText="Policy Request Details"
            closeButtonAriaLabel="Close"
          >
            <div style={{ padding: '16px 0' }}>
              {/* Status & Priority Header */}
              <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 20 }}>
                <span style={{
                  background: `${this.getRequestStatusColor(selectedPolicyRequest.Status)}15`,
                  color: this.getRequestStatusColor(selectedPolicyRequest.Status),
                  padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
                }}>
                  {selectedPolicyRequest.Status === 'InProgress' ? 'In Progress' : selectedPolicyRequest.Status}
                </span>
                <span style={{
                  background: `${this.getPriorityColor(selectedPolicyRequest.Priority)}15`,
                  color: this.getPriorityColor(selectedPolicyRequest.Priority),
                  padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
                }}>
                  {selectedPolicyRequest.Priority} Priority
                </span>
              </Stack>

              {/* Title */}
              <Text variant="xLarge" style={{ fontWeight: 700, display: 'block', marginBottom: 16 }}>{selectedPolicyRequest.Title}</Text>

              {/* Section: Request Details */}
              <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Request Information</Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Requested By:</Text>
                    <Text>{selectedPolicyRequest.RequestedBy} ({selectedPolicyRequest.RequestedByDepartment})</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Email:</Text>
                    <Text>{selectedPolicyRequest.RequestedByEmail}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Category:</Text>
                    <Text>{selectedPolicyRequest.PolicyCategory}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Type:</Text>
                    <Text>{selectedPolicyRequest.PolicyType}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Submitted:</Text>
                    <Text>{new Date(selectedPolicyRequest.Created).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
                  </Stack>
                </Stack>
              </div>

              {/* Section: Business Justification */}
              <div style={{ background: '#fffbeb', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #f59e0b' }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Business Justification</Text>
                <Text style={{ lineHeight: '1.6' }}>{selectedPolicyRequest.BusinessJustification}</Text>
              </div>

              {/* Section: Regulatory Driver */}
              {selectedPolicyRequest.RegulatoryDriver && (
                <div style={{ background: '#fef2f2', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #ef4444' }}>
                  <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Regulatory / Compliance Driver</Text>
                  <Text>{selectedPolicyRequest.RegulatoryDriver}</Text>
                </div>
              )}

              {/* Section: Policy Requirements */}
              <div style={{ background: '#f0fdfa', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #0d9488' }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Policy Requirements</Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Target Audience:</Text>
                    <Text>{selectedPolicyRequest.TargetAudience}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Desired Effective Date:</Text>
                    <Text>{new Date(selectedPolicyRequest.DesiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Read Timeframe:</Text>
                    <Text>{selectedPolicyRequest.ReadTimeframeDays} days</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Acknowledgement:</Text>
                    <Text style={{ color: selectedPolicyRequest.RequiresAcknowledgement ? '#107c10' : '#605e5c' }}>
                      {selectedPolicyRequest.RequiresAcknowledgement ? 'Yes' : 'No'}
                    </Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Quiz:</Text>
                    <Text style={{ color: selectedPolicyRequest.RequiresQuiz ? '#8764b8' : '#605e5c' }}>
                      {selectedPolicyRequest.RequiresQuiz ? 'Yes' : 'No'}
                    </Text>
                  </Stack>
                </Stack>
              </div>

              {/* Section: Additional Notes */}
              {selectedPolicyRequest.AdditionalNotes && (
                <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
                  <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Additional Notes</Text>
                  <Text style={{ lineHeight: '1.6', fontStyle: 'italic' }}>{selectedPolicyRequest.AdditionalNotes}</Text>
                </div>
              )}

              {/* Section: Assignment */}
              <div style={{ background: '#f3eefc', borderRadius: 8, padding: 16, marginBottom: 20 }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Assignment</Text>
                {selectedPolicyRequest.AssignedAuthor ? (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <div style={{ width: 36, height: 36, borderRadius: '50%', background: '#8764b8', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 14 }}>
                      {selectedPolicyRequest.AssignedAuthor.split(' ').map(n => n[0]).join('').slice(0, 2)}
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>{selectedPolicyRequest.AssignedAuthor}</Text>
                      <Text variant="small" style={{ display: 'block', color: '#605e5c' }}>{selectedPolicyRequest.AssignedAuthorEmail}</Text>
                    </div>
                  </Stack>
                ) : (
                  <Text style={{ color: '#a19f9d', fontStyle: 'italic' }}>Not yet assigned — click "Accept & Start" below</Text>
                )}
              </div>

              {/* Action Buttons */}
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {(selectedPolicyRequest.Status === 'New' || selectedPolicyRequest.Status === 'Assigned') && (
                  <PrimaryButton
                    text="Accept & Start Drafting"
                    iconProps={{ iconName: 'Play' }}
                    onClick={() => {
                      const updated = { ...selectedPolicyRequest, Status: 'InProgress' as const, AssignedAuthor: context.pageContext.user.displayName, AssignedAuthorEmail: context.pageContext.user.email };
                      onSetState({
                        selectedPolicyRequest: updated,
                        policyRequests: policyRequests.map(r => r.Id === updated.Id ? updated : r)
                      });
                      // Also open the Policy Builder with request data pre-populated
                      onCreatePolicyFromRequest(updated);
                    }}
                  />
                )}
                {selectedPolicyRequest.Status === 'InProgress' && (
                  <PrimaryButton
                    text="Mark as Draft Ready"
                    iconProps={{ iconName: 'CheckMark' }}
                    styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                    onClick={() => {
                      const updated = { ...selectedPolicyRequest, Status: 'Draft Ready' as const };
                      onSetState({
                        selectedPolicyRequest: updated,
                        policyRequests: policyRequests.map(r => r.Id === updated.Id ? updated : r)
                      });
                    }}
                  />
                )}
                <DefaultButton
                  text="Create Policy from Request"
                  iconProps={{ iconName: 'PageAdd' }}
                  onClick={() => onCreatePolicyFromRequest(selectedPolicyRequest)}
                />
                <DefaultButton
                  text="Close"
                  onClick={() => onSetState({ showPolicyRequestDetailPanel: false, selectedPolicyRequest: null })}
                />
              </Stack>
            </div>
          </Panel>
        )}
      </>
    );
  }
}
