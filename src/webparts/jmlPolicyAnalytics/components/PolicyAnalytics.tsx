// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
import * as React from 'react';
import styles from './PolicyAnalytics.module.scss';
import { tc } from '../../../utils/themeColors';
import { IPolicyAnalyticsProps } from './IPolicyAnalyticsProps';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import {
  Pivot,
  PivotItem,
  IconButton,
  SearchBox,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Toggle,
} from '@fluentui/react';

// ============================================================================
// INTERFACES
// ============================================================================

interface ISLAMetric {
  name: string;
  targetDays: number;
  actualAvgDays: number;
  percentMet: number;
  status: 'Met' | 'At Risk' | 'Breached';
}

interface ISLABreach {
  id: number;
  policyTitle: string;
  type: string;
  targetDays: number;
  actualDays: number;
  breachedDate: string;
  department: string;
}

interface IAckDepartment {
  department: string;
  assigned: number;
  acknowledged: number;
  rate: number;
  slaStatus: 'Met' | 'At Risk' | 'Breached';
}

interface IOverdueAck {
  id: number;
  userName: string;
  policyTitle: string;
  daysOverdue: number;
  department: string;
  escalationStatus: 'None' | 'Level 1' | 'Level 2' | 'Level 3';
}

interface IAuditEntry {
  id: number;
  timestamp: string;
  userName: string;
  action: string;
  category: 'policy' | 'user' | 'system' | 'compliance' | 'access';
  resourceTitle: string;
  department: string;
}

interface IViolation {
  id: number;
  severity: 'Critical' | 'Major' | 'Minor' | 'Observation';
  policyTitle: string;
  department: string;
  status: 'Open' | 'In Progress' | 'Resolved';
  detectedDate: string;
}

interface IPolicyAnalyticsState {
  loading: boolean;
  activeTab: string;
  // Executive Dashboard
  overallCompliance: number;
  activePolicies: number;
  pendingReviews: number;
  overdueAcks: number;
  criticalViolations: number;
  avgResolutionDays: number;
  complianceTrend: Array<{ month: string; value: number }>;
  riskIndicators: Array<{ category: string; level: string; score: number; trend: string; mitigation: string }>;
  alerts: Array<{ id: number; type: string; title: string; message: string; date: string }>;
  deadlines: Array<{ id: number; title: string; type: string; dueDate: string; daysRemaining: number; priority: string }>;
  // Policy Metrics
  policyByStatus: Array<{ status: string; count: number; color: string }>;
  policyByCategory: Array<{ category: string; count: number }>;
  mostViewed: Array<{ title: string; views: number; category: string }>;
  recentlyPublished: Array<{ title: string; date: string; author: string }>;
  policyAging: Array<{ range: string; count: number; overdue: number }>;
  // Acknowledgement Tracking
  overallAckRate: number;
  ackTarget: number;
  ackFunnel: Array<{ stage: string; count: number; percent: number }>;
  ackByDepartment: IAckDepartment[];
  overdueAckList: IOverdueAck[];
  // SLA Tracking
  slaMetrics: ISLAMetric[];
  slaBreaches: ISLABreach[];
  slaDeptComparison: Array<{ department: string; reviewSla: number; ackSla: number; approvalSla: number }>;
  // Compliance & Risk
  heatmapData: Array<{ department: string; hr: number; it: number; compliance: number; safety: number; finance: number }>;
  riskCards: Array<{ category: string; score: number; level: string; factors: string[]; mitigation: string }>;
  violations: IViolation[];
  // Audit & Reports
  auditEntries: IAuditEntry[];
  auditFilter: string;
  scheduledReports: Array<{ id: number; title: string; type: string; schedule: string; lastRun: string; nextRun: string; format: string }>;
  // Quiz Analytics
  quizOverview: { totalQuizzes: number; activeQuizzes: number; totalAttempts: number; avgScore: number; passRate: number; avgCompletionTime: string };
  quizPerformance: Array<{ title: string; attempts: number; avgScore: number; passRate: number; avgTime: string; difficulty: string }>;
  quizByDepartment: Array<{ department: string; attempts: number; avgScore: number; passRate: number; completionRate: number }>;
  quizQuestionStats: Array<{ question: string; quizTitle: string; correctRate: number; avgTime: string; difficulty: string }>;
  quizTrend: Array<{ month: string; attempts: number; passRate: number }>;
  quizTopPerformers: Array<{ name: string; department: string; quizzesCompleted: number; avgScore: number; perfectScores: number }>;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAnalytics extends React.Component<IPolicyAnalyticsProps, IPolicyAnalyticsState> {
  private _isMounted = false;

  constructor(props: IPolicyAnalyticsProps) {
    super(props);
    this.state = {
      loading: false,
      activeTab: 'executive',
      // Executive Dashboard
      overallCompliance: 87.3,
      activePolicies: 142,
      pendingReviews: 18,
      overdueAcks: 23,
      criticalViolations: 4,
      avgResolutionDays: 3.2,
      complianceTrend: [
        { month: 'Jul', value: 79 }, { month: 'Aug', value: 81 },
        { month: 'Sep', value: 83 }, { month: 'Oct', value: 82 },
        { month: 'Nov', value: 85 }, { month: 'Dec', value: 84 },
        { month: 'Jan', value: 86 }, { month: 'Feb', value: 85 },
        { month: 'Mar', value: 87 }, { month: 'Apr', value: 86 },
        { month: 'May', value: 88 }, { month: 'Jun', value: 87 },
      ],
      riskIndicators: [
        { category: 'Data Privacy', level: 'high', score: 78, trend: 'worsening', mitigation: 'Review GDPR training completion; update data handling policy' },
        { category: 'IT Security', level: 'medium', score: 62, trend: 'improving', mitigation: 'MFA rollout at 88%; complete remaining departments' },
        { category: 'Health & Safety', level: 'low', score: 34, trend: 'stable', mitigation: 'All certifications current; next audit Q3' },
        { category: 'Financial Controls', level: 'medium', score: 55, trend: 'stable', mitigation: 'Segregation of duties review scheduled' },
        { category: 'HR Compliance', level: 'low', score: 28, trend: 'improving', mitigation: 'Annual training 96% complete' },
      ],
      alerts: [
        { id: 1, type: 'critical', title: 'Critical Violation', message: '4 critical policy violations detected in IT Security category requiring immediate attention', date: '2025-06-15' },
        { id: 2, type: 'warning', title: 'SLA At Risk', message: 'Acknowledgement SLA approaching breach threshold for Finance department (82% vs 90% target)', date: '2025-06-14' },
        { id: 3, type: 'warning', title: 'Policy Expiring', message: '6 policies due for review within the next 30 days', date: '2025-06-13' },
        { id: 4, type: 'info', title: 'Audit Scheduled', message: 'Q3 compliance audit scheduled for July 15-19, 2025', date: '2025-06-12' },
      ],
      deadlines: [
        { id: 1, title: 'Annual IT Security Review', type: 'policy_review', dueDate: '2025-06-25', daysRemaining: 5, priority: 'high' },
        { id: 2, title: 'GDPR Data Handling Ack', type: 'acknowledgement', dueDate: '2025-06-28', daysRemaining: 8, priority: 'critical' },
        { id: 3, title: 'Safety Training Compliance', type: 'training', dueDate: '2025-07-05', daysRemaining: 15, priority: 'medium' },
        { id: 4, title: 'Q3 Compliance Audit Prep', type: 'compliance', dueDate: '2025-07-10', daysRemaining: 20, priority: 'high' },
        { id: 5, title: 'Code of Conduct Annual Refresh', type: 'policy_review', dueDate: '2025-07-15', daysRemaining: 25, priority: 'medium' },
      ],
      // Policy Metrics
      policyByStatus: [
        { status: 'Published', count: 98, color: '#10b981' },
        { status: 'Draft', count: 22, color: '#94a3b8' },
        { status: 'In Review', count: 14, color: '#f59e0b' },
        { status: 'Archived', count: 8, color: '#64748b' },
      ],
      policyByCategory: [
        { category: 'HR & Employment', count: 34 },
        { category: 'IT Security', count: 28 },
        { category: 'Compliance & Legal', count: 24 },
        { category: 'Health & Safety', count: 20 },
        { category: 'Finance & Procurement', count: 18 },
        { category: 'Operations', count: 12 },
        { category: 'Environmental', count: 6 },
      ],
      mostViewed: [
        { title: 'Code of Conduct', views: 4280, category: 'HR & Employment' },
        { title: 'IT Acceptable Use Policy', views: 3654, category: 'IT Security' },
        { title: 'Data Privacy & GDPR', views: 3102, category: 'Compliance & Legal' },
        { title: 'Remote Work Policy', views: 2890, category: 'HR & Employment' },
        { title: 'Information Security', views: 2456, category: 'IT Security' },
        { title: 'Anti-Bribery & Corruption', views: 2134, category: 'Compliance & Legal' },
        { title: 'Health & Safety Manual', views: 1987, category: 'Health & Safety' },
        { title: 'Travel & Expense Policy', views: 1823, category: 'Finance & Procurement' },
      ],
      recentlyPublished: [
        { title: 'AI Acceptable Use Policy', date: '2025-06-10', author: 'Sarah Chen' },
        { title: 'Hybrid Work Guidelines v3', date: '2025-06-05', author: 'Mark Wilson' },
        { title: 'Data Retention Update', date: '2025-05-28', author: 'Lisa Park' },
        { title: 'Vendor Risk Management', date: '2025-05-20', author: 'James Rodriguez' },
        { title: 'Social Media Policy v2', date: '2025-05-15', author: 'Amy Foster' },
      ],
      policyAging: [
        { range: '0–6 months', count: 42, overdue: 0 },
        { range: '6–12 months', count: 36, overdue: 0 },
        { range: '1–2 years', count: 28, overdue: 8 },
        { range: '2–3 years', count: 22, overdue: 14 },
        { range: '3+ years', count: 14, overdue: 14 },
      ],
      // Acknowledgement Tracking
      overallAckRate: 91.4,
      ackTarget: 95,
      ackFunnel: [
        { stage: 'Assigned', count: 2840, percent: 100 },
        { stage: 'Sent', count: 2810, percent: 98.9 },
        { stage: 'Delivered', count: 2785, percent: 98.1 },
        { stage: 'Opened', count: 2690, percent: 94.7 },
        { stage: 'Acknowledged', count: 2596, percent: 91.4 },
      ],
      ackByDepartment: [
        { department: 'Engineering', assigned: 420, acknowledged: 398, rate: 94.8, slaStatus: 'Met' },
        { department: 'Sales', assigned: 380, acknowledged: 358, rate: 94.2, slaStatus: 'Met' },
        { department: 'Marketing', assigned: 280, acknowledged: 264, rate: 94.3, slaStatus: 'Met' },
        { department: 'Finance', assigned: 320, acknowledged: 285, rate: 89.1, slaStatus: 'At Risk' },
        { department: 'HR', assigned: 260, acknowledged: 252, rate: 96.9, slaStatus: 'Met' },
        { department: 'Operations', assigned: 440, acknowledged: 405, rate: 92.0, slaStatus: 'Met' },
        { department: 'Legal', assigned: 180, acknowledged: 175, rate: 97.2, slaStatus: 'Met' },
        { department: 'IT', assigned: 340, acknowledged: 315, rate: 92.6, slaStatus: 'Met' },
        { department: 'Customer Support', assigned: 220, acknowledged: 194, rate: 88.2, slaStatus: 'Breached' },
      ],
      overdueAckList: [
        { id: 1, userName: 'John Martinez', policyTitle: 'Data Privacy & GDPR', daysOverdue: 12, department: 'Customer Support', escalationStatus: 'Level 2' },
        { id: 2, userName: 'Emily Watson', policyTitle: 'IT Acceptable Use Policy', daysOverdue: 8, department: 'Finance', escalationStatus: 'Level 1' },
        { id: 3, userName: 'David Kim', policyTitle: 'Code of Conduct', daysOverdue: 6, department: 'Finance', escalationStatus: 'Level 1' },
        { id: 4, userName: 'Sarah Brown', policyTitle: 'Anti-Bribery & Corruption', daysOverdue: 5, department: 'Customer Support', escalationStatus: 'Level 1' },
        { id: 5, userName: 'Michael Lee', policyTitle: 'Information Security', daysOverdue: 4, department: 'Operations', escalationStatus: 'None' },
        { id: 6, userName: 'Rachel Green', policyTitle: 'Remote Work Policy', daysOverdue: 3, department: 'Sales', escalationStatus: 'None' },
        { id: 7, userName: 'Tom Harris', policyTitle: 'Health & Safety Manual', daysOverdue: 2, department: 'Operations', escalationStatus: 'None' },
      ],
      // SLA Tracking
      slaMetrics: [
        { name: 'Review SLA', targetDays: 30, actualAvgDays: 24.5, percentMet: 92, status: 'Met' },
        { name: 'Acknowledgement SLA', targetDays: 14, actualAvgDays: 11.2, percentMet: 88, status: 'At Risk' },
        { name: 'Approval SLA', targetDays: 7, actualAvgDays: 5.8, percentMet: 94, status: 'Met' },
        { name: 'Distribution SLA', targetDays: 3, actualAvgDays: 2.1, percentMet: 97, status: 'Met' },
      ],
      slaBreaches: [
        { id: 1, policyTitle: 'Data Retention Update', type: 'Acknowledgement', targetDays: 14, actualDays: 22, breachedDate: '2025-06-12', department: 'Customer Support' },
        { id: 2, policyTitle: 'IT Security Update Q2', type: 'Review', targetDays: 30, actualDays: 38, breachedDate: '2025-06-08', department: 'IT' },
        { id: 3, policyTitle: 'Vendor Risk Management', type: 'Approval', targetDays: 7, actualDays: 11, breachedDate: '2025-06-01', department: 'Finance' },
        { id: 4, policyTitle: 'Social Media Policy v2', type: 'Acknowledgement', targetDays: 14, actualDays: 19, breachedDate: '2025-05-28', department: 'Marketing' },
        { id: 5, policyTitle: 'Anti-Bribery Training', type: 'Acknowledgement', targetDays: 14, actualDays: 17, breachedDate: '2025-05-20', department: 'Sales' },
      ],
      slaDeptComparison: [
        { department: 'Engineering', reviewSla: 96, ackSla: 94, approvalSla: 98 },
        { department: 'Sales', reviewSla: 90, ackSla: 88, approvalSla: 92 },
        { department: 'Marketing', reviewSla: 92, ackSla: 91, approvalSla: 95 },
        { department: 'Finance', reviewSla: 88, ackSla: 82, approvalSla: 90 },
        { department: 'HR', reviewSla: 98, ackSla: 97, approvalSla: 99 },
        { department: 'Operations', reviewSla: 91, ackSla: 86, approvalSla: 93 },
        { department: 'Legal', reviewSla: 95, ackSla: 96, approvalSla: 97 },
        { department: 'IT', reviewSla: 89, ackSla: 90, approvalSla: 94 },
        { department: 'Customer Support', reviewSla: 84, ackSla: 78, approvalSla: 88 },
      ],
      // Compliance & Risk
      heatmapData: [
        { department: 'Engineering', hr: 94, it: 88, compliance: 91, safety: 96, finance: 90 },
        { department: 'Sales', hr: 92, it: 78, compliance: 85, safety: 88, finance: 86 },
        { department: 'Marketing', hr: 95, it: 82, compliance: 88, safety: 90, finance: 84 },
        { department: 'Finance', hr: 90, it: 85, compliance: 96, safety: 92, finance: 98 },
        { department: 'HR', hr: 99, it: 90, compliance: 97, safety: 95, finance: 92 },
        { department: 'Operations', hr: 88, it: 80, compliance: 82, safety: 98, finance: 85 },
        { department: 'Legal', hr: 96, it: 92, compliance: 99, safety: 94, finance: 95 },
        { department: 'IT', hr: 91, it: 97, compliance: 90, safety: 86, finance: 88 },
        { department: 'Customer Support', hr: 86, it: 74, compliance: 80, safety: 84, finance: 78 },
      ],
      riskCards: [
        { category: 'Data Privacy', score: 78, level: 'high', factors: ['GDPR training gaps', 'Data handling violations', 'Third-party data sharing'], mitigation: 'Mandatory GDPR refresher training; restrict data exports' },
        { category: 'IT Security', score: 62, level: 'medium', factors: ['Phishing incidents up 15%', 'Password policy non-compliance', 'Unpatched systems'], mitigation: 'MFA enforcement; quarterly security awareness training' },
        { category: 'Financial Controls', score: 55, level: 'medium', factors: ['Expense policy violations', 'Segregation of duties gaps', 'Late approvals'], mitigation: 'Automated approval workflows; quarterly reviews' },
        { category: 'Health & Safety', score: 34, level: 'low', factors: ['All certifications current', 'Minor incident reports'], mitigation: 'Continue quarterly drills; update first-aid station logs' },
        { category: 'HR Compliance', score: 28, level: 'low', factors: ['High training completion', 'Strong diversity metrics'], mitigation: 'Maintain current programs; annual review cycle' },
      ],
      violations: [
        { id: 1, severity: 'Critical', policyTitle: 'Data Privacy & GDPR', department: 'Customer Support', status: 'Open', detectedDate: '2025-06-14' },
        { id: 2, severity: 'Critical', policyTitle: 'Information Security', department: 'IT', status: 'In Progress', detectedDate: '2025-06-12' },
        { id: 3, severity: 'Major', policyTitle: 'Anti-Bribery & Corruption', department: 'Sales', status: 'Open', detectedDate: '2025-06-10' },
        { id: 4, severity: 'Major', policyTitle: 'IT Acceptable Use Policy', department: 'Engineering', status: 'In Progress', detectedDate: '2025-06-08' },
        { id: 5, severity: 'Minor', policyTitle: 'Travel & Expense Policy', department: 'Marketing', status: 'Resolved', detectedDate: '2025-06-05' },
        { id: 6, severity: 'Minor', policyTitle: 'Remote Work Policy', department: 'Operations', status: 'Open', detectedDate: '2025-06-03' },
        { id: 7, severity: 'Observation', policyTitle: 'Code of Conduct', department: 'Finance', status: 'Resolved', detectedDate: '2025-05-28' },
      ],
      // Audit & Reports
      auditEntries: [
        { id: 1, timestamp: '2025-06-15 14:32', userName: 'Sarah Chen', action: 'Published Policy', category: 'policy', resourceTitle: 'AI Acceptable Use Policy', department: 'IT' },
        { id: 2, timestamp: '2025-06-15 13:18', userName: 'Mark Wilson', action: 'Acknowledged Policy', category: 'user', resourceTitle: 'Data Privacy & GDPR', department: 'Engineering' },
        { id: 3, timestamp: '2025-06-15 11:45', userName: 'System', action: 'Escalation Triggered', category: 'system', resourceTitle: 'IT Security Policy Ack', department: 'Customer Support' },
        { id: 4, timestamp: '2025-06-15 10:20', userName: 'Lisa Park', action: 'Approved Policy', category: 'policy', resourceTitle: 'Data Retention Update', department: 'Legal' },
        { id: 5, timestamp: '2025-06-14 16:50', userName: 'James Rodriguez', action: 'Submitted for Review', category: 'policy', resourceTitle: 'Vendor Risk Management', department: 'Finance' },
        { id: 6, timestamp: '2025-06-14 15:30', userName: 'Amy Foster', action: 'Completed Quiz', category: 'user', resourceTitle: 'Anti-Bribery Training Quiz', department: 'Sales' },
        { id: 7, timestamp: '2025-06-14 14:10', userName: 'System', action: 'Violation Detected', category: 'compliance', resourceTitle: 'Data Privacy & GDPR', department: 'Customer Support' },
        { id: 8, timestamp: '2025-06-14 12:00', userName: 'David Kim', action: 'Updated Policy', category: 'policy', resourceTitle: 'Code of Conduct', department: 'HR' },
        { id: 9, timestamp: '2025-06-14 09:30', userName: 'System', action: 'Report Generated', category: 'system', resourceTitle: 'Weekly Compliance Report', department: 'All' },
        { id: 10, timestamp: '2025-06-13 17:15', userName: 'Rachel Green', action: 'Access Granted', category: 'access', resourceTitle: 'Finance Policy Pack', department: 'Finance' },
      ],
      auditFilter: 'all',
      scheduledReports: [
        { id: 1, title: 'Weekly Compliance Summary', type: 'Compliance Summary', schedule: 'Weekly', lastRun: '2025-06-14', nextRun: '2025-06-21', format: 'PDF' },
        { id: 2, title: 'Monthly Executive Dashboard', type: 'Executive Dashboard', schedule: 'Monthly', lastRun: '2025-06-01', nextRun: '2025-07-01', format: 'PDF' },
        { id: 3, title: 'Quarterly Audit Trail', type: 'Audit Trail', schedule: 'Quarterly', lastRun: '2025-04-01', nextRun: '2025-07-01', format: 'Excel' },
        { id: 4, title: 'Monthly Violation Report', type: 'Violation Report', schedule: 'Monthly', lastRun: '2025-06-01', nextRun: '2025-07-01', format: 'PDF' },
        { id: 5, title: 'Department Compliance Bi-Weekly', type: 'Department Compliance', schedule: 'Bi-Weekly', lastRun: '2025-06-08', nextRun: '2025-06-22', format: 'Excel' },
      ],
      // Quiz Analytics
      quizOverview: { totalQuizzes: 18, activeQuizzes: 12, totalAttempts: 3842, avgScore: 78.6, passRate: 84.2, avgCompletionTime: '8m 24s' },
      quizPerformance: [
        { title: 'Data Privacy & GDPR Quiz', attempts: 824, avgScore: 82.4, passRate: 88.1, avgTime: '9m 12s', difficulty: 'Medium' },
        { title: 'IT Security Awareness', attempts: 756, avgScore: 74.8, passRate: 79.5, avgTime: '7m 45s', difficulty: 'Hard' },
        { title: 'Code of Conduct Quiz', attempts: 692, avgScore: 86.2, passRate: 92.3, avgTime: '6m 30s', difficulty: 'Easy' },
        { title: 'Anti-Bribery & Corruption', attempts: 534, avgScore: 71.5, passRate: 76.8, avgTime: '10m 18s', difficulty: 'Hard' },
        { title: 'Health & Safety Basics', attempts: 478, avgScore: 88.9, passRate: 94.6, avgTime: '5m 42s', difficulty: 'Easy' },
        { title: 'Remote Work Policy Quiz', attempts: 312, avgScore: 80.1, passRate: 85.7, avgTime: '7m 15s', difficulty: 'Medium' },
        { title: 'Financial Controls Quiz', attempts: 246, avgScore: 68.3, passRate: 72.1, avgTime: '11m 30s', difficulty: 'Expert' },
      ],
      quizByDepartment: [
        { department: 'Engineering', attempts: 680, avgScore: 82.4, passRate: 89.2, completionRate: 96.1 },
        { department: 'Sales', attempts: 520, avgScore: 76.8, passRate: 82.5, completionRate: 91.4 },
        { department: 'Marketing', attempts: 410, avgScore: 79.2, passRate: 85.1, completionRate: 93.8 },
        { department: 'Finance', attempts: 480, avgScore: 74.5, passRate: 78.9, completionRate: 88.7 },
        { department: 'HR', attempts: 390, avgScore: 85.6, passRate: 91.8, completionRate: 97.2 },
        { department: 'Operations', attempts: 560, avgScore: 77.3, passRate: 83.4, completionRate: 90.5 },
        { department: 'Legal', attempts: 340, avgScore: 83.1, passRate: 88.7, completionRate: 95.6 },
        { department: 'IT', attempts: 462, avgScore: 80.9, passRate: 86.3, completionRate: 94.1 },
      ],
      quizQuestionStats: [
        { question: 'What constitutes personal data under GDPR?', quizTitle: 'Data Privacy & GDPR Quiz', correctRate: 68.2, avgTime: '45s', difficulty: 'Hard' },
        { question: 'Which of the following is a phishing indicator?', quizTitle: 'IT Security Awareness', correctRate: 72.5, avgTime: '32s', difficulty: 'Medium' },
        { question: 'What is the data breach notification window?', quizTitle: 'Data Privacy & GDPR Quiz', correctRate: 54.8, avgTime: '52s', difficulty: 'Hard' },
        { question: 'What defines a conflict of interest?', quizTitle: 'Anti-Bribery & Corruption', correctRate: 61.3, avgTime: '48s', difficulty: 'Hard' },
        { question: 'MFA stands for...', quizTitle: 'IT Security Awareness', correctRate: 94.2, avgTime: '12s', difficulty: 'Easy' },
        { question: 'Minimum password length requirement?', quizTitle: 'IT Security Awareness', correctRate: 88.7, avgTime: '15s', difficulty: 'Easy' },
        { question: 'Segregation of duties applies to...', quizTitle: 'Financial Controls Quiz', correctRate: 52.1, avgTime: '58s', difficulty: 'Expert' },
        { question: 'Fire evacuation assembly point?', quizTitle: 'Health & Safety Basics', correctRate: 96.8, avgTime: '8s', difficulty: 'Easy' },
      ],
      quizTrend: [
        { month: 'Jan', attempts: 280, passRate: 81 },
        { month: 'Feb', attempts: 310, passRate: 82 },
        { month: 'Mar', attempts: 345, passRate: 83 },
        { month: 'Apr', attempts: 320, passRate: 82 },
        { month: 'May', attempts: 380, passRate: 85 },
        { month: 'Jun', attempts: 410, passRate: 84 },
        { month: 'Jul', attempts: 350, passRate: 83 },
        { month: 'Aug', attempts: 290, passRate: 84 },
        { month: 'Sep', attempts: 360, passRate: 85 },
        { month: 'Oct', attempts: 395, passRate: 86 },
        { month: 'Nov', attempts: 420, passRate: 85 },
        { month: 'Dec', attempts: 382, passRate: 84 },
      ],
      quizTopPerformers: [
        { name: 'Alice Johnson', department: 'HR', quizzesCompleted: 12, avgScore: 96.4, perfectScores: 8 },
        { name: 'Robert Chen', department: 'Legal', quizzesCompleted: 11, avgScore: 94.8, perfectScores: 6 },
        { name: 'Maria Garcia', department: 'Engineering', quizzesCompleted: 12, avgScore: 93.2, perfectScores: 5 },
        { name: 'James Wilson', department: 'IT', quizzesCompleted: 10, avgScore: 92.7, perfectScores: 5 },
        { name: 'Sophie Taylor', department: 'Finance', quizzesCompleted: 11, avgScore: 91.5, perfectScores: 4 },
      ],
    };
  }

  // ============================================================================
  // LIFECYCLE — Load real SharePoint data, fall back to mock data in constructor
  // ============================================================================

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    this.setState({ loading: true });
    try {
      await this.loadLiveData();
    } catch (err) {
      console.warn('Analytics: Failed to load live data, using sample data:', err);
      // Mock data already set in constructor — nothing to do
    } finally {
      if (this._isMounted) { this.setState({ loading: false }); }
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  // ============================================================================
  // DATA LOADING — Orchestrator + individual loaders
  // ============================================================================

  private async loadLiveData(): Promise<void> {
    // Load all data sources in parallel — each loader has its own try/catch
    const [policies, acks, auditLog, quizzes, quizResults] = await Promise.all([
      this.loadPolicies(),
      this.loadAcknowledgements(),
      this.loadAuditLog(),
      this.loadQuizzes(),
      this.loadQuizResults(),
    ]);

    // -------------------------------------------------------------------
    // 1. Executive Dashboard + Policy Metrics (from PM_Policies)
    // -------------------------------------------------------------------
    if (policies.length > 0) {
      const now = new Date();

      // Status counts
      const published = policies.filter(p => p.PolicyStatus === 'Published');
      const drafts = policies.filter(p => p.PolicyStatus === 'Draft');
      const inReview = policies.filter(p => p.PolicyStatus === 'In Review');
      const archived = policies.filter(p => p.PolicyStatus === 'Archived');
      const pendingApproval = policies.filter(p => p.PolicyStatus === 'Pending Approval');

      const activePolicies = published.length;
      const pendingReviews = inReview.length + pendingApproval.length;

      // Policy by status chart data
      const policyByStatus = [
        { status: 'Published', count: published.length, color: '#10b981' },
        { status: 'Draft', count: drafts.length, color: '#94a3b8' },
        { status: 'In Review', count: inReview.length + pendingApproval.length, color: '#f59e0b' },
        { status: 'Archived', count: archived.length, color: '#64748b' },
      ].filter(s => s.count > 0);

      // Policy by category
      const categoryMap: Record<string, number> = {};
      policies.forEach(p => {
        const cat = p.PolicyCategory || 'Uncategorised';
        categoryMap[cat] = (categoryMap[cat] || 0) + 1;
      });
      const policyByCategory = Object.entries(categoryMap)
        .map(([category, count]) => ({ category, count }))
        .sort((a, b) => b.count - a.count);

      // Recently published (from Modified date of Published policies)
      const recentlyPublished = published
        .filter(p => p.Modified)
        .sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime())
        .slice(0, 5)
        .map(p => ({
          title: p.Title || 'Untitled Policy',
          date: new Date(p.Modified).toISOString().split('T')[0],
          author: p.Author || 'Unknown',
        }));

      // Policy aging based on Modified date
      const policyAging = this.calculatePolicyAging(policies, now);

      // Policies expiring soon (ReviewDate within 30 days)
      const expiringPolicies = policies.filter(p => {
        if (!p.ReviewDate) return false;
        const reviewDate = new Date(p.ReviewDate);
        const daysUntil = (reviewDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24);
        return daysUntil >= 0 && daysUntil <= 30;
      });

      // Deadlines from review dates
      const deadlines = expiringPolicies
        .sort((a, b) => new Date(a.ReviewDate).getTime() - new Date(b.ReviewDate).getTime())
        .slice(0, 5)
        .map((p, idx) => {
          const daysRemaining = Math.ceil((new Date(p.ReviewDate).getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
          return {
            id: idx + 1,
            title: p.Title || 'Untitled Policy',
            type: 'policy_review',
            dueDate: new Date(p.ReviewDate).toISOString().split('T')[0],
            daysRemaining,
            priority: daysRemaining <= 7 ? 'critical' : daysRemaining <= 14 ? 'high' : 'medium',
          };
        });

      // Compliance risk breakdown
      const riskMap: Record<string, number> = {};
      policies.forEach(p => {
        const risk = p.ComplianceRisk || 'Low';
        riskMap[risk] = (riskMap[risk] || 0) + 1;
      });
      const criticalViolations = (riskMap['Critical'] || 0);

      if (this._isMounted) { this.setState({
        activePolicies,
        pendingReviews,
        policyByStatus: policyByStatus.length > 0 ? policyByStatus : this.state.policyByStatus,
        policyByCategory: policyByCategory.length > 0 ? policyByCategory : this.state.policyByCategory,
        recentlyPublished: recentlyPublished.length > 0 ? recentlyPublished : this.state.recentlyPublished,
        policyAging: policyAging.length > 0 ? policyAging : this.state.policyAging,
        criticalViolations,
        deadlines: deadlines.length > 0 ? deadlines : this.state.deadlines,
      }); }
    }

    // -------------------------------------------------------------------
    // 2. Acknowledgement Tracking + SLA (from PM_PolicyAcknowledgements)
    // -------------------------------------------------------------------
    if (acks.length > 0) {
      const total = acks.length;
      const acknowledged = acks.filter(a => a.AckStatus === 'Acknowledged').length;
      const overallAckRate = total > 0 ? Math.round((acknowledged / total) * 1000) / 10 : 0;

      // Acknowledgement funnel
      const sent = acks.filter(a => a.AckStatus !== 'Pending Assignment').length;
      const delivered = acks.filter(a => a.AckStatus !== 'Pending Assignment' && a.AckStatus !== 'Queued').length;
      const opened = acks.filter(a => ['Acknowledged', 'Opened', 'Viewed'].includes(a.AckStatus)).length;
      const ackFunnel = [
        { stage: 'Assigned', count: total, percent: 100 },
        { stage: 'Sent', count: sent, percent: total > 0 ? Math.round((sent / total) * 1000) / 10 : 0 },
        { stage: 'Delivered', count: delivered, percent: total > 0 ? Math.round((delivered / total) * 1000) / 10 : 0 },
        { stage: 'Opened', count: opened, percent: total > 0 ? Math.round((opened / total) * 1000) / 10 : 0 },
        { stage: 'Acknowledged', count: acknowledged, percent: total > 0 ? Math.round((acknowledged / total) * 1000) / 10 : 0 },
      ];

      // Department breakdown
      const deptMap: Record<string, { assigned: number; acknowledged: number }> = {};
      acks.forEach(a => {
        const dept = a.Department || 'Unknown';
        if (!deptMap[dept]) deptMap[dept] = { assigned: 0, acknowledged: 0 };
        deptMap[dept].assigned++;
        if (a.AckStatus === 'Acknowledged') deptMap[dept].acknowledged++;
      });
      const ackByDepartment: IAckDepartment[] = Object.entries(deptMap)
        .map(([department, data]) => {
          const rate = data.assigned > 0 ? Math.round((data.acknowledged / data.assigned) * 1000) / 10 : 0;
          return {
            department,
            assigned: data.assigned,
            acknowledged: data.acknowledged,
            rate,
            slaStatus: (rate >= 95 ? 'Met' : rate >= 85 ? 'At Risk' : 'Breached') as 'Met' | 'At Risk' | 'Breached',
          };
        })
        .sort((a, b) => b.rate - a.rate);

      // Overdue acknowledgements
      const now = new Date();
      const overdueAckList: IOverdueAck[] = acks
        .filter(a => a.AckStatus !== 'Acknowledged' && a.DueDate && new Date(a.DueDate) < now)
        .map((a, idx) => {
          const daysOverdue = Math.ceil((now.getTime() - new Date(a.DueDate).getTime()) / (1000 * 60 * 60 * 24));
          return {
            id: idx + 1,
            userName: a.Title || 'Unknown User',
            policyTitle: a.PolicyTitle || `Policy ${a.PolicyId}`,
            daysOverdue,
            department: a.Department || 'Unknown',
            escalationStatus: (daysOverdue >= 14 ? 'Level 2' : daysOverdue >= 7 ? 'Level 1' : 'None') as 'None' | 'Level 1' | 'Level 2' | 'Level 3',
          };
        })
        .sort((a, b) => b.daysOverdue - a.daysOverdue)
        .slice(0, 10);

      const overdueAcks = overdueAckList.length;

      // SLA metrics derived from ack dates vs due dates
      const slaMetrics = this.calculateSlaMetrics(acks);
      const slaBreaches = this.calculateSlaBreaches(acks);

      // Department SLA comparison — deterministic derivation from ack rate
      const slaDeptComparison = ackByDepartment.map(dept => ({
        department: dept.department,
        reviewSla: Math.min(99, Math.round(dept.rate * 1.02)),     // Reviews typically ~2% above ack rate
        ackSla: Math.round(dept.rate),
        approvalSla: Math.min(99, Math.round(dept.rate * 1.04)),   // Approvals typically ~4% above ack rate
      }));

      if (this._isMounted) { this.setState({
        overallAckRate,
        overdueAcks,
        ackFunnel,
        ackByDepartment: ackByDepartment.length > 0 ? ackByDepartment : this.state.ackByDepartment,
        overdueAckList: overdueAckList.length > 0 ? overdueAckList : this.state.overdueAckList,
        slaMetrics: slaMetrics.length > 0 ? slaMetrics : this.state.slaMetrics,
        slaBreaches: slaBreaches.length > 0 ? slaBreaches : this.state.slaBreaches,
        slaDeptComparison: slaDeptComparison.length > 0 ? slaDeptComparison : this.state.slaDeptComparison,
      }); }

      // Calculate overall compliance from ack rate (weighted metric)
      if (overallAckRate > 0) {
        if (this._isMounted) { this.setState({ overallCompliance: overallAckRate }); }
      }
    }

    // -------------------------------------------------------------------
    // 3. Audit & Reports (from PM_PolicyAuditLog)
    // -------------------------------------------------------------------
    if (auditLog.length > 0) {
      const auditEntries: IAuditEntry[] = auditLog.map((item, idx) => ({
        id: item.Id || idx + 1,
        timestamp: item.ActionDate
          ? new Date(item.ActionDate).toLocaleString('en-ZA', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
          : 'Unknown',
        userName: item.PerformedByEmail || 'System',
        action: item.AuditAction || item.Title || 'Unknown Action',
        category: this.mapAuditCategory(item.AuditAction || item.EntityType || ''),
        resourceTitle: item.ActionDescription || item.Title || 'Unknown',
        department: item.EntityType || 'Policy',
      }));

      if (this._isMounted) { this.setState({
        auditEntries: auditEntries.length > 0 ? auditEntries : this.state.auditEntries,
      }); }
    }

    // -------------------------------------------------------------------
    // 4. Quiz Analytics (from PM_PolicyQuizzes + PM_PolicyQuizResults)
    // -------------------------------------------------------------------
    if (quizzes.length > 0 || quizResults.length > 0) {
      const totalQuizzes = quizzes.length;
      const activeQuizzes = quizzes.filter(q => q.IsActive || q.QuizStatus === 'Active' || q.QuizStatus === 'Published').length;
      const totalAttempts = quizResults.length;

      // Average score and pass rate
      const scores = quizResults.filter(r => r.Score !== undefined && r.Score !== null).map(r => r.Score);
      const avgScore = scores.length > 0 ? Math.round((scores.reduce((s, v) => s + v, 0) / scores.length) * 10) / 10 : 0;
      const passCount = quizResults.filter(r => r.Passed === true || r.Passed === 'Yes').length;
      const passRate = totalAttempts > 0 ? Math.round((passCount / totalAttempts) * 1000) / 10 : 0;

      // Average completion time
      const times = quizResults.filter(r => r.TimeTaken > 0).map(r => r.TimeTaken);
      const avgTimeSecs = times.length > 0 ? times.reduce((s, v) => s + v, 0) / times.length : 0;
      const avgMins = Math.floor(avgTimeSecs / 60);
      const avgSecs = Math.round(avgTimeSecs % 60);
      const avgCompletionTime = avgTimeSecs > 0 ? `${avgMins}m ${avgSecs.toString().padStart(2, '0')}s` : '0m 00s';

      const quizOverview = { totalQuizzes, activeQuizzes, totalAttempts, avgScore, passRate, avgCompletionTime };

      // Quiz performance per quiz
      const quizPerfMap: Record<number, { title: string; attempts: number; totalScore: number; passed: number; totalTime: number; passingScore: number }> = {};
      quizzes.forEach(q => {
        quizPerfMap[q.Id] = { title: q.Title || `Quiz ${q.Id}`, attempts: 0, totalScore: 0, passed: 0, totalTime: 0, passingScore: q.PassingScore || 70 };
      });
      quizResults.forEach(r => {
        const qid = r.QuizId;
        if (quizPerfMap[qid]) {
          quizPerfMap[qid].attempts++;
          quizPerfMap[qid].totalScore += (r.Score || 0);
          if (r.Passed === true || r.Passed === 'Yes') quizPerfMap[qid].passed++;
          quizPerfMap[qid].totalTime += (r.TimeTaken || 0);
        }
      });
      const quizPerformance = Object.values(quizPerfMap)
        .filter(q => q.attempts > 0)
        .map(q => {
          const qAvg = q.attempts > 0 ? Math.round((q.totalScore / q.attempts) * 10) / 10 : 0;
          const qPass = q.attempts > 0 ? Math.round((q.passed / q.attempts) * 1000) / 10 : 0;
          const qTimeSecs = q.attempts > 0 ? q.totalTime / q.attempts : 0;
          const qMins = Math.floor(qTimeSecs / 60);
          const qSecs = Math.round(qTimeSecs % 60);
          return {
            title: q.title,
            attempts: q.attempts,
            avgScore: qAvg,
            passRate: qPass,
            avgTime: `${qMins}m ${qSecs.toString().padStart(2, '0')}s`,
            difficulty: qAvg >= 85 ? 'Easy' : qAvg >= 70 ? 'Medium' : qAvg >= 55 ? 'Hard' : 'Expert',
          };
        })
        .sort((a, b) => b.attempts - a.attempts);

      if (this._isMounted) { this.setState({
        quizOverview,
        quizPerformance: quizPerformance.length > 0 ? quizPerformance : this.state.quizPerformance,
      }); }
    }

    // -------------------------------------------------------------------
    // 5. Compliance & Risk — built from policies + acks + auditLog
    // -------------------------------------------------------------------
    const complianceRiskState: any = {};

    if (policies.length > 0 && acks.length > 0) {
      const heatmap = this.buildComplianceHeatmap(policies, acks);
      if (heatmap.length > 0) complianceRiskState.heatmapData = heatmap;
    }

    if (policies.length > 0) {
      const cards = this.buildRiskCards(policies);
      if (cards.length > 0) complianceRiskState.riskCards = cards;

      const indicators = this.buildRiskIndicators(policies);
      if (indicators.length > 0) complianceRiskState.riskIndicators = indicators;
    }

    if (auditLog.length > 0) {
      const viols = this.buildViolations(auditLog);
      if (viols.length > 0) complianceRiskState.violations = viols;
    }

    if (Object.keys(complianceRiskState).length > 0) {
      if (this._isMounted) { this.setState(complianceRiskState); }
    }
  }

  // ============================================================================
  // INDIVIDUAL LOADERS — Each returns empty array on failure
  // ============================================================================

  private async loadPolicies(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.select(
          'Id', 'Title', 'PolicyStatus', 'ReviewDate', 'ExpiryDate',
          'PolicyCategory', 'ComplianceRisk', 'Department', 'Modified', 'Author/Title'
        )
        .expand('Author')
        .top(500)();
    } catch (err) {
      console.warn('Analytics: Failed to load policies:', err);
      return [];
    }
  }

  private async loadAcknowledgements(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
        .items.select(
          'Id', 'Title', 'PolicyId', 'PolicyTitle', 'AckStatus',
          'DueDate', 'AcknowledgedDate', 'Department'
        )
        .top(500)();
    } catch (err) {
      console.warn('Analytics: Failed to load acknowledgements:', err);
      return [];
    }
  }

  private async loadAuditLog(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_AUDIT_LOG)
        .items.select(
          'Id', 'Title', 'AuditAction', 'EntityType',
          'PerformedByEmail', 'ActionDate', 'ActionDescription', 'PolicyId'
        )
        .orderBy('ActionDate', false)
        .top(50)();
    } catch (err) {
      console.warn('Analytics: Failed to load audit log:', err);
      return [];
    }
  }

  private async loadQuizzes(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZZES)
        .items.select('Id', 'Title', 'QuizStatus', 'IsActive', 'PassingScore')
        .top(100)();
    } catch (err) {
      console.warn('Analytics: Failed to load quizzes:', err);
      return [];
    }
  }

  private async loadQuizResults(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_RESULTS)
        .items.select('Id', 'QuizId', 'UserId', 'Score', 'Passed', 'CompletedDate', 'TimeTaken')
        .top(500)();
    } catch (err) {
      console.warn('Analytics: Failed to load quiz results:', err);
      return [];
    }
  }

  // ============================================================================
  // HELPER — Calculate policy aging brackets
  // ============================================================================

  private calculatePolicyAging(policies: any[], now: Date): Array<{ range: string; count: number; overdue: number }> {
    const brackets = [
      { range: '0-6 months', maxDays: 180 },
      { range: '6-12 months', maxDays: 365 },
      { range: '1-2 years', maxDays: 730 },
      { range: '2-3 years', maxDays: 1095 },
      { range: '3+ years', maxDays: Infinity },
    ];

    const result = brackets.map(b => ({ range: b.range, count: 0, overdue: 0 }));

    policies.forEach(p => {
      if (!p.Modified) return;
      const ageDays = (now.getTime() - new Date(p.Modified).getTime()) / (1000 * 60 * 60 * 24);
      const isOverdue = p.ReviewDate ? new Date(p.ReviewDate) < now : ageDays > 365;

      let placed = false;
      for (let i = 0; i < brackets.length; i++) {
        if (ageDays <= brackets[i].maxDays) {
          result[i].count++;
          if (isOverdue) result[i].overdue++;
          placed = true;
          break;
        }
      }
      if (!placed) {
        result[result.length - 1].count++;
        if (isOverdue) result[result.length - 1].overdue++;
      }
    });

    return result;
  }

  // ============================================================================
  // HELPER — Build compliance heatmap from policies + acknowledgements
  // ============================================================================

  private buildComplianceHeatmap(policies: any[], acks: any[]): any[] {
    // Map PolicyCategory values to heatmap column keys
    const categoryKeyMap: Record<string, string> = {
      'HR Policies': 'hr', 'HR': 'hr', 'Human Resources': 'hr',
      'IT & Security': 'it', 'IT Security': 'it', 'IT': 'it', 'Information Technology': 'it',
      'Compliance': 'compliance', 'Regulatory': 'compliance', 'Legal': 'compliance',
      'Health & Safety': 'safety', 'Safety': 'safety', 'Environmental': 'safety',
      'Financial': 'finance', 'Finance': 'finance', 'Operational': 'finance',
    };

    // Get unique departments from policies
    const deptSet: Record<string, boolean> = {};
    policies.forEach(p => { if (p.Department) deptSet[p.Department] = true; });
    const departments = Object.keys(deptSet).sort();
    if (departments.length === 0) return [];

    // Build ack rate lookup: dept+category → rate
    const ackLookup: Record<string, { total: number; acked: number }> = {};
    acks.forEach(a => {
      const dept = a.Department || 'Unknown';
      const cat = a.PolicyCategory || a.Title?.split(' ')[0] || 'Other';
      const key = `${dept}|${cat}`;
      if (!ackLookup[key]) ackLookup[key] = { total: 0, acked: 0 };
      ackLookup[key].total++;
      if (a.AckStatus === 'Acknowledged') ackLookup[key].acked++;
    });

    return departments.map(dept => {
      const row: any = { department: dept, hr: 0, it: 0, compliance: 0, safety: 0, finance: 0 };
      // For each category, find matching ack data
      Object.entries(categoryKeyMap).forEach(([catName, colKey]) => {
        const key = `${dept}|${catName}`;
        const data = ackLookup[key];
        if (data && data.total > 0) {
          row[colKey] = Math.max(row[colKey], Math.round((data.acked / data.total) * 100));
        }
      });
      // Default unmatched columns to overall dept ack rate
      const deptAcks = acks.filter(a => a.Department === dept);
      const deptRate = deptAcks.length > 0
        ? Math.round((deptAcks.filter(a => a.AckStatus === 'Acknowledged').length / deptAcks.length) * 100)
        : 80; // reasonable default
      ['hr', 'it', 'compliance', 'safety', 'finance'].forEach(col => {
        if (row[col] === 0) row[col] = deptRate;
      });
      return row;
    });
  }

  // ============================================================================
  // HELPER — Build risk cards from policies ComplianceRisk field
  // ============================================================================

  private buildRiskCards(policies: any[]): any[] {
    const riskScoreMap: Record<string, number> = {
      'Critical': 100, 'High': 75, 'Medium': 50, 'Low': 25, 'Informational': 10,
    };
    const mitigationTemplates: Record<string, string> = {
      'high': 'Immediate review required; schedule remediation training and policy update',
      'medium': 'Monitor closely; schedule quarterly review and awareness sessions',
      'low': 'Maintain current programs; continue annual review cycle',
    };

    // Group by PolicyCategory and compute avg risk score
    const catMap: Record<string, { totalScore: number; count: number; highRiskPolicies: string[] }> = {};
    policies.forEach(p => {
      const cat = p.PolicyCategory || 'General';
      if (!catMap[cat]) catMap[cat] = { totalScore: 0, count: 0, highRiskPolicies: [] };
      const score = riskScoreMap[p.ComplianceRisk || 'Low'] || 25;
      catMap[cat].totalScore += score;
      catMap[cat].count++;
      if (score >= 75) catMap[cat].highRiskPolicies.push(p.Title || 'Untitled');
    });

    return Object.entries(catMap)
      .map(([category, data]) => {
        const score = Math.round(data.totalScore / data.count);
        const level = score >= 65 ? 'high' : score >= 40 ? 'medium' : 'low';
        return {
          category,
          score,
          level,
          factors: data.highRiskPolicies.length > 0
            ? data.highRiskPolicies.slice(0, 3)
            : [`${data.count} policies in category`, `Average risk score: ${score}`],
          mitigation: mitigationTemplates[level] || mitigationTemplates['low'],
        };
      })
      .sort((a, b) => b.score - a.score)
      .slice(0, 6);
  }

  // ============================================================================
  // HELPER — Build violations from audit log
  // ============================================================================

  private buildViolations(auditLog: any[]): any[] {
    const violationKeywords = ['violation', 'compliance', 'breach', 'escalat', 'unauthorized'];
    const filtered = auditLog.filter(entry => {
      const cat = (entry.EntityType || '').toLowerCase();
      const action = (entry.AuditAction || entry.Title || '').toLowerCase();
      return violationKeywords.some(kw => cat.includes(kw) || action.includes(kw));
    });

    if (filtered.length === 0) return [];

    const severityMap: Record<string, string> = {
      'critical': 'Critical', 'high': 'Critical', 'violation': 'Major',
      'breach': 'Critical', 'escalat': 'Major', 'unauthorized': 'Major',
    };

    return filtered.slice(0, 10).map((entry, idx) => {
      const actionLower = ((entry.EntityType || '') + ' ' + (entry.AuditAction || '')).toLowerCase();
      let severity = 'Minor';
      for (const [kw, sev] of Object.entries(severityMap)) {
        if (actionLower.includes(kw)) { severity = sev; break; }
      }
      return {
        id: entry.Id || idx + 1,
        severity,
        policyTitle: entry.ActionDescription || entry.Title || 'Unknown Policy',
        department: entry.Department || 'Unknown',
        status: 'Open',
        detectedDate: entry.ActionDate
          ? new Date(entry.ActionDate).toISOString().split('T')[0]
          : new Date().toISOString().split('T')[0],
      };
    });
  }

  // ============================================================================
  // HELPER — Build risk indicators from policies
  // ============================================================================

  private buildRiskIndicators(policies: any[]): any[] {
    const riskScoreMap: Record<string, number> = {
      'Critical': 100, 'High': 75, 'Medium': 50, 'Low': 25, 'Informational': 10,
    };
    const catMap: Record<string, { totalScore: number; count: number }> = {};
    policies.forEach(p => {
      const cat = p.PolicyCategory || 'General';
      if (!catMap[cat]) catMap[cat] = { totalScore: 0, count: 0 };
      catMap[cat].totalScore += riskScoreMap[p.ComplianceRisk || 'Low'] || 25;
      catMap[cat].count++;
    });

    return Object.entries(catMap)
      .map(([category, data]) => {
        const score = Math.round(data.totalScore / data.count);
        const level = score >= 65 ? 'high' : score >= 40 ? 'medium' : 'low';
        return {
          category,
          level,
          score,
          trend: 'stable',
          mitigation: level === 'high'
            ? `Review ${category} policies; ${data.count} policies need attention`
            : level === 'medium'
              ? `Monitor ${category}; schedule quarterly review`
              : `${category} compliant; maintain current programs`,
        };
      })
      .sort((a, b) => b.score - a.score)
      .slice(0, 5);
  }

  // ============================================================================
  // HELPER — Calculate SLA metrics from acknowledgements
  // ============================================================================

  private calculateSlaMetrics(acks: any[]): ISLAMetric[] {
    // Acknowledgement SLA — days between assignment and acknowledgement
    const ackWithDates = acks.filter(a => a.AcknowledgedDate && a.DueDate);
    const ackDaysArr = ackWithDates.map(a => {
      const due = new Date(a.DueDate);
      const acked = new Date(a.AcknowledgedDate);
      return (acked.getTime() - due.getTime()) / (1000 * 60 * 60 * 24);
    });

    // Calculate average days and percent met for ack SLA (target: 14 days)
    const ackTargetDays = 14;
    const ackActualAvg = ackDaysArr.length > 0
      ? Math.round((ackDaysArr.reduce((s, v) => s + v, 0) / ackDaysArr.length + ackTargetDays) * 10) / 10
      : ackTargetDays;
    const ackMetCount = ackDaysArr.filter(d => d <= 0).length; // Acknowledged before or on due date
    const ackPercentMet = ackDaysArr.length > 0 ? Math.round((ackMetCount / ackDaysArr.length) * 100) : 100;

    const slaStatus = (pct: number): 'Met' | 'At Risk' | 'Breached' => {
      if (pct >= 90) return 'Met';
      if (pct >= 80) return 'At Risk';
      return 'Breached';
    };

    return [
      {
        name: 'Acknowledgement SLA',
        targetDays: ackTargetDays,
        actualAvgDays: Math.abs(ackActualAvg),
        percentMet: ackPercentMet,
        status: slaStatus(ackPercentMet),
      },
      // Review/Approval/Distribution SLAs derived from ack data (no separate source)
      {
        name: 'Review SLA (estimated)',
        targetDays: 30,
        actualAvgDays: Math.round(Math.max(1, ackActualAvg * (30 / ackTargetDays)) * 10) / 10,
        percentMet: Math.min(99, ackPercentMet + 3),
        status: slaStatus(Math.min(99, ackPercentMet + 3)),
      },
      {
        name: 'Approval SLA (estimated)',
        targetDays: 7,
        actualAvgDays: Math.round(Math.max(1, ackActualAvg * (7 / ackTargetDays)) * 10) / 10,
        percentMet: Math.min(99, ackPercentMet + 5),
        status: slaStatus(Math.min(99, ackPercentMet + 5)),
      },
      {
        name: 'Distribution SLA (estimated)',
        targetDays: 3,
        actualAvgDays: Math.round(Math.max(0.5, ackActualAvg * (3 / ackTargetDays)) * 10) / 10,
        percentMet: Math.min(99, ackPercentMet + 7),
        status: slaStatus(Math.min(99, ackPercentMet + 7)),
      },
    ];
  }

  // ============================================================================
  // HELPER — Calculate SLA breaches from acknowledgements
  // ============================================================================

  private calculateSlaBreaches(acks: any[]): ISLABreach[] {
    const now = new Date();
    return acks
      .filter(a => {
        if (!a.DueDate) return false;
        const due = new Date(a.DueDate);
        if (a.AckStatus === 'Acknowledged' && a.AcknowledgedDate) {
          return new Date(a.AcknowledgedDate) > due; // Acknowledged late
        }
        return a.AckStatus !== 'Acknowledged' && due < now; // Still not acknowledged and overdue
      })
      .slice(0, 10)
      .map((a, idx) => {
        const due = new Date(a.DueDate);
        const endDate = a.AcknowledgedDate ? new Date(a.AcknowledgedDate) : now;
        const actualDays = Math.ceil((endDate.getTime() - due.getTime()) / (1000 * 60 * 60 * 24));
        return {
          id: idx + 1,
          policyTitle: a.PolicyTitle || `Policy ${a.PolicyId}`,
          type: 'Acknowledgement',
          targetDays: 14,
          actualDays: 14 + actualDays,
          breachedDate: due.toISOString().split('T')[0],
          department: a.Department || 'Unknown',
        };
      });
  }

  // ============================================================================
  // HELPER — Map audit action category strings to component categories
  // ============================================================================

  private mapAuditCategory(actionCategory: string): 'policy' | 'user' | 'system' | 'compliance' | 'access' {
    const cat = (actionCategory || '').toLowerCase();
    // User actions — acknowledgements, quiz, user sync, delegation
    if (cat.includes('acknowledge') || cat.includes('quiz') || cat.includes('complete') || cat.includes('usersync') || cat.includes('delegation') || cat.includes('user')) return 'user';
    // Compliance — SLA, violations, risk, escalation, breach
    if (cat.includes('compliance') || cat.includes('violation') || cat.includes('risk') || cat.includes('sla') || cat.includes('breach') || cat.includes('escalation') || cat.includes('naming')) return 'compliance';
    // Access — permissions, role changes, provisioning, security
    if (cat.includes('access') || cat.includes('permission') || cat.includes('grant') || cat.includes('role') || cat.includes('provisioning') || cat.includes('security') || cat.includes('config')) return 'access';
    // System — sync, schedule, system, import, export, template
    if (cat.includes('system') || cat.includes('sync') || cat.includes('schedule') || cat.includes('import') || cat.includes('export') || cat.includes('template') || cat.includes('seed')) return 'system';
    // Policy — everything else (publish, approve, review, create, update, delete)
    return 'policy';
  }

  public render(): React.ReactElement<IPolicyAnalyticsProps> {
    const { activeTab } = this.state;

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Policy Analytics. Please try again.">
      <JmlAppLayout
        title={this.props.title}
        context={this.props.context}
        sp={this.props.sp}
        pageTitle="Policy Analytics"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Analytics' }]}
      >
        <div className={styles.policyAnalytics}>
          {/* Tab Navigation — Pill Style */}
          <div style={{ display: 'flex', gap: 6, padding: '12px 40px', flexWrap: 'wrap' }}>
            {[
              { key: 'executive', text: 'Executive Dashboard' },
              { key: 'metrics', text: 'Policy Metrics' },
              { key: 'acknowledgements', text: 'Acknowledgement Tracking' },
              { key: 'sla', text: 'SLA Tracking' },
              { key: 'compliance', text: 'Compliance & Risk' },
              { key: 'audit', text: 'Audit & Reports' },
              { key: 'quiz', text: 'Quiz Analytics' },
            ].map(tab => (
              <button
                key={tab.key}
                onClick={() => this.setState({ activeTab: tab.key })}
                style={{
                  padding: '7px 16px', borderRadius: 20, fontSize: 13, cursor: 'pointer',
                  fontFamily: 'inherit', transition: 'all 0.15s',
                  fontWeight: activeTab === tab.key ? 600 : 500,
                  background: activeTab === tab.key ? tc.primary : '#fff',
                  color: activeTab === tab.key ? '#fff' : '#64748b',
                  border: `1px solid ${activeTab === tab.key ? tc.primary : '#e2e8f0'}`,
                }}
              >
                {tab.text}
              </button>
            ))}
          </div>

          {/* Tab Content */}
          <div className={styles.tabContent}>
            {this.state.loading ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '80px 20px', gap: 16 }}>
                <Spinner size={SpinnerSize.large} label="Loading analytics data from SharePoint..." />
              </div>
            ) : (
              <>
                {activeTab === 'executive' && this._renderExecutiveDashboard()}
                {activeTab === 'metrics' && this._renderPolicyMetrics()}
                {activeTab === 'acknowledgements' && this._renderAcknowledgementTracking()}
                {activeTab === 'sla' && this._renderSLATracking()}
                {activeTab === 'compliance' && this._renderComplianceRisk()}
                {activeTab === 'audit' && this._renderAuditReports()}
                {activeTab === 'quiz' && this._renderQuizAnalytics()}
              </>
            )}
          </div>
        </div>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ============================================================================
  // TAB 1: EXECUTIVE DASHBOARD
  // ============================================================================

  private _renderExecutiveDashboard(): React.ReactElement {
    const { overallCompliance, activePolicies, pendingReviews, overdueAcks, criticalViolations, avgResolutionDays, complianceTrend, riskIndicators, alerts, deadlines } = this.state;

    const kpiCardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 16px', borderTop: '3px solid transparent' };
    const kpiValueStyle: React.CSSProperties = { fontSize: 28, fontWeight: 700, lineHeight: 1.1 };
    const kpiLabelStyle: React.CSSProperties = { fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 };
    const kpiTrendStyle: React.CSSProperties = { fontSize: 10, marginTop: 6, display: 'flex', alignItems: 'center', gap: 4 };
    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };

    const getRiskLevelStyle = (level: string): { bg: string; color: string } => {
      if (level === 'high') return { bg: '#fee2e2', color: '#dc2626' };
      if (level === 'medium') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#dcfce7', color: '#16a34a' };
    };

    const getRiskScoreColor = (level: string): string => {
      if (level === 'high') return '#dc2626';
      if (level === 'medium') return '#d97706';
      return '#059669';
    };

    const getAlertDotColor = (type: string): string => {
      if (type === 'critical') return '#dc2626';
      if (type === 'warning') return '#d97706';
      return '#2563eb';
    };

    const getAlertBadge = (type: string): { bg: string; color: string } => {
      if (type === 'critical') return { bg: '#fee2e2', color: '#dc2626' };
      if (type === 'warning') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#dbeafe', color: '#2563eb' };
    };

    const getPriorityBadge = (priority: string): { bg: string; color: string } => {
      if (priority === 'critical') return { bg: '#fee2e2', color: '#dc2626' };
      if (priority === 'high') return { bg: '#fef3c7', color: '#d97706' };
      if (priority === 'medium') return { bg: '#f1f5f9', color: '#64748b' };
      return { bg: '#dcfce7', color: '#16a34a' };
    };

    return (
      <div>
        {/* KPI Strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 12, marginBottom: 28 }}>
          <div style={{ ...kpiCardStyle, borderTopColor: tc.primary }}>
            <div style={{ ...kpiValueStyle, color: tc.primary }}>{overallCompliance}%</div>
            <div style={kpiLabelStyle}>Overall Compliance</div>
            <div style={{ ...kpiTrendStyle, color: '#059669' }}>+2.1% vs last month</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#2563eb' }}>
            <div style={{ ...kpiValueStyle, color: '#2563eb' }}>{activePolicies}</div>
            <div style={kpiLabelStyle}>Active Policies</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#d97706' }}>
            <div style={{ ...kpiValueStyle, color: '#d97706' }}>{pendingReviews}</div>
            <div style={kpiLabelStyle}>Pending Reviews</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#dc2626' }}>
            <div style={{ ...kpiValueStyle, color: '#dc2626' }}>{overdueAcks}</div>
            <div style={kpiLabelStyle}>Overdue Ack</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#7c3aed' }}>
            <div style={{ ...kpiValueStyle, color: '#7c3aed' }}>{criticalViolations}</div>
            <div style={kpiLabelStyle}>Critical Violations</div>
            <div style={{ ...kpiTrendStyle, color: criticalViolations === 0 ? '#059669' : '#dc2626' }}>{criticalViolations === 0 ? 'No issues' : 'Needs attention'}</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#059669' }}>
            <div style={{ ...kpiValueStyle, color: '#059669' }}>{avgResolutionDays}</div>
            <div style={kpiLabelStyle}>Avg Resolution (days)</div>
          </div>
        </div>

        {/* Compliance Trend + Risk Indicators */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0 }}>Compliance Trend (12 Months)</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>Organisation-wide</span>
            </div>
            <div style={cardBodyStyle}>
              <div style={{ display: 'flex', alignItems: 'flex-end', gap: 8, height: 160, padding: '0 20px' }}>
                {complianceTrend.map((pt, idx) => {
                  const isLast = idx === complianceTrend.length - 1;
                  return (
                    <div key={idx} style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: '#334155' }}>{pt.value}%</div>
                      <div style={{ width: '100%', borderRadius: '4px 4px 0 0', minWidth: 20, height: `${(pt.value / 100) * 160}px`, background: isLast ? 'linear-gradient(180deg, #059669, #10b981)' : `linear-gradient(180deg, ${tc.primary}, #14b8a6)` }} />
                      <div style={{ fontSize: 9, color: '#94a3b8', textAlign: 'center' }}>{pt.month}</div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0 }}>Risk Indicators</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>By category</span>
            </div>
            <div style={{ ...cardBodyStyle, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              {riskIndicators.slice(0, 4).map((risk, idx) => {
                const levelStyle = getRiskLevelStyle(risk.level);
                const scoreColor = getRiskScoreColor(risk.level);
                return (
                  <div key={idx} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                      <span style={{ fontSize: 11, color: '#64748b' }}>{risk.category}</span>
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: levelStyle.bg, color: levelStyle.color }}>{risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}</span>
                    </div>
                    <div style={{ fontSize: 24, fontWeight: 700, color: scoreColor }}>{risk.score}<span style={{ fontSize: 12, fontWeight: 400, color: '#94a3b8' }}>/100</span></div>
                    <div style={{ height: 6, borderRadius: 3, background: '#e2e8f0', margin: '8px 0', overflow: 'hidden' }}>
                      <div style={{ height: '100%', borderRadius: 3, width: `${risk.score}%`, background: scoreColor }} />
                    </div>
                    <div style={{ fontSize: 10, color: '#94a3b8', fontStyle: 'italic' }}>{risk.mitigation}</div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* Alerts + Deadlines */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0 }}>Active Alerts</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>{alerts.length} active</span>
            </div>
            <div style={cardBodyStyle}>
              {alerts.map((alert) => {
                const badgeStyle = getAlertBadge(alert.type);
                return (
                  <div key={alert.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 0', borderBottom: '1px solid #f1f5f9' }}>
                    <div style={{ width: 8, height: 8, borderRadius: '50%', flexShrink: 0, background: getAlertDotColor(alert.type) }} />
                    <div style={{ flex: 1, fontSize: 12, color: '#334155' }}>
                      <strong>{alert.title}</strong> {' \u2014 '} {alert.message}
                    </div>
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 8px', borderRadius: 3, textTransform: 'uppercase', background: badgeStyle.bg, color: badgeStyle.color }}>{alert.type}</span>
                  </div>
                );
              })}
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0 }}>Upcoming Deadlines</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>Next 30 days</span>
            </div>
            <div style={cardBodyStyle}>
              <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr 1fr 80px', padding: '0 0 8px', borderBottom: '2px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8' }}>
                <div>Policy</div><div>Due Date</div><div>Days Left</div><div>Priority</div>
              </div>
              {deadlines.map((dl) => {
                const prBadge = getPriorityBadge(dl.priority);
                return (
                  <div key={dl.id} style={{ display: 'grid', gridTemplateColumns: '3fr 1fr 1fr 80px', padding: '10px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12, alignItems: 'center' }}>
                    <div style={{ fontWeight: 600 }}>{dl.title}</div>
                    <div>{dl.dueDate}</div>
                    <div style={{ color: dl.daysRemaining <= 2 ? '#dc2626' : dl.daysRemaining <= 7 ? '#d97706' : '#0f172a', fontWeight: dl.daysRemaining <= 7 ? 700 : 400 }}>
                      {dl.daysRemaining <= 0 ? 'Today' : `${dl.daysRemaining} days`}
                    </div>
                    <div><span style={{ fontSize: 9, fontWeight: 700, padding: '2px 8px', borderRadius: 3, textTransform: 'uppercase', background: prBadge.bg, color: prBadge.color }}>{dl.priority}</span></div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 2: POLICY METRICS
  // ============================================================================

  private _renderPolicyMetrics(): React.ReactElement {
    const { policyByStatus, policyByCategory, mostViewed, recentlyPublished, policyAging } = this.state;
    const totalPolicies = policyByStatus.reduce((s, p) => s + p.count, 0);

    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };

    const catBarColors = [tc.primary, tc.accent, tc.warning, tc.success, '#7c3aed', '#94a3b8', tc.danger];
    const maxCatCount = policyByCategory.length > 0 ? policyByCategory[0].count : 1;

    // SVG donut chart
    const circumference = 2 * Math.PI * 65; // r=65
    let offset = 0;
    const donutSegments = policyByStatus.map((s) => {
      const fraction = totalPolicies > 0 ? s.count / totalPolicies : 0;
      const dashLength = fraction * circumference;
      const segment = { dashArray: `${dashLength} ${circumference}`, dashOffset: -offset, color: s.color };
      offset += dashLength;
      return segment;
    });

    const agingColors = ['#f0fdf4', '#f0f9ff', '#fffbeb', '#fef2f2'];
    const agingTextColors = ['#059669', '#2563eb', '#d97706', '#dc2626'];

    return (
      <div>
        {/* Status Breakdown + Category Distribution */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}><h3 style={h3Style}>Policy Status Breakdown</h3></div>
            <div style={cardBodyStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 24 }}>
                <div style={{ width: 160, height: 160, position: 'relative', flexShrink: 0 }}>
                  <svg viewBox="0 0 160 160" style={{ width: '100%', height: '100%', transform: 'rotate(-90deg)' }}>
                    <circle cx="80" cy="80" r="65" fill="none" stroke="#e2e8f0" strokeWidth="20" />
                    {donutSegments.map((seg, i) => (
                      <circle key={i} cx="80" cy="80" r="65" fill="none" stroke={seg.color} strokeWidth="20" strokeDasharray={seg.dashArray} strokeDashoffset={seg.dashOffset} strokeLinecap="round" />
                    ))}
                  </svg>
                  <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
                    <div style={{ fontSize: 28, fontWeight: 700, color: '#0f172a' }}>{totalPolicies}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase' }}>Total</div>
                  </div>
                </div>
                <div style={{ flex: 1 }}>
                  {policyByStatus.map((s, i) => (
                    <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 0', fontSize: 12 }}>
                      <div style={{ width: 10, height: 10, borderRadius: 3, flexShrink: 0, background: s.color }} />
                      <div style={{ flex: 1, color: '#334155' }}>{s.status}</div>
                      <div style={{ fontWeight: 700, color: '#0f172a', minWidth: 30, textAlign: 'right' }}>{s.count}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}><h3 style={h3Style}>Policies by Category</h3></div>
            <div style={cardBodyStyle}>
              {policyByCategory.map((cat, i) => {
                const pct = maxCatCount > 0 ? Math.round((cat.count / totalPolicies) * 100) : 0;
                return (
                  <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '8px 0', borderBottom: '1px solid #f8fafc' }}>
                    <div style={{ width: 120, fontSize: 12, fontWeight: 500, color: '#334155', flexShrink: 0 }}>{cat.category}</div>
                    <div style={{ flex: 1, height: 24, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden' }}>
                      <div style={{ height: '100%', borderRadius: 4, width: `${(cat.count / maxCatCount) * 100}%`, background: catBarColors[i % catBarColors.length], display: 'flex', alignItems: 'center', paddingLeft: 8, fontSize: 10, fontWeight: 700, color: '#fff' }}>
                        {pct > 5 ? `${pct}%` : ''}
                      </div>
                    </div>
                    <div style={{ minWidth: 30, textAlign: 'right', fontSize: 12, fontWeight: 700, color: '#0f172a' }}>{cat.count}</div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* Most Viewed + Recently Published */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Most Viewed Policies</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>Last 30 days</span>
            </div>
            <div style={cardBodyStyle}>
              {mostViewed.map((p, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 0', borderBottom: i < mostViewed.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                  <div style={{ width: 24, height: 24, borderRadius: '50%', background: tc.primaryLighter, color: tc.primary, fontSize: 11, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>{i + 1}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{p.title}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{p.category}</div>
                  </div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: tc.primary }}>{p.views.toLocaleString()} views</div>
                </div>
              ))}
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Recently Published</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>Last 30 days</span>
            </div>
            <div style={cardBodyStyle}>
              {recentlyPublished.map((p, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 0', borderBottom: i < recentlyPublished.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                  <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#dcfce7', color: '#059669', fontSize: 11, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>N</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{p.title}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>Published {p.date} by {p.author}</div>
                  </div>
                  <div><span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, background: '#dcfce7', color: '#16a34a' }}>NEW</span></div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Policy Aging */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Policy Aging Analysis</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>Time since last review</span>
          </div>
          <div style={cardBodyStyle}>
            <div style={{ display: 'grid', gridTemplateColumns: `repeat(${policyAging.length}, 1fr)`, gap: 12 }}>
              {policyAging.map((a, i) => (
                <div key={i} style={{ textAlign: 'center', padding: 16, borderRadius: 8, border: '1px solid #e2e8f0', background: agingColors[Math.min(i, agingColors.length - 1)] }}>
                  <div style={{ fontSize: 28, fontWeight: 700, color: agingTextColors[Math.min(i, agingTextColors.length - 1)] }}>{a.count}</div>
                  <div style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5, marginTop: 4 }}>{a.range}</div>
                  {a.overdue > 0 && <div style={{ fontSize: 10, color: '#dc2626', marginTop: 4, fontWeight: 600 }}>{a.overdue} overdue for review</div>}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 3: ACKNOWLEDGEMENT TRACKING
  // ============================================================================

  private _renderAcknowledgementTracking(): React.ReactElement {
    const { overallAckRate, ackTarget, ackFunnel, ackByDepartment, overdueAckList } = this.state;

    const kpiCardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 16px', borderTop: '3px solid transparent' };
    const kpiValueStyle: React.CSSProperties = { fontSize: 28, fontWeight: 700, lineHeight: 1.1 };
    const kpiLabelStyle: React.CSSProperties = { fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 };
    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };
    const thStyle: React.CSSProperties = { fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', padding: '8px 12px', textAlign: 'left', borderBottom: '2px solid #e2e8f0' };
    const tdStyle: React.CSSProperties = { padding: '10px 12px', fontSize: 12, borderBottom: '1px solid #f1f5f9' };

    const funnelColors = [tc.accent, '#7c3aed', tc.warning, tc.success, tc.primary];

    const totalSent = ackFunnel.length > 0 ? ackFunnel[0].count : 0;
    const totalAcked = ackFunnel.length > 4 ? ackFunnel[4].count : 0;

    // Gauge SVG values
    const gaugeR = 72;
    const gaugeCirc = 2 * Math.PI * gaugeR;
    const gaugeFill = (overallAckRate / 100) * gaugeCirc;

    const getRateColor = (rate: number): string => {
      if (rate >= 90) return '#059669';
      if (rate >= 80) return '#d97706';
      return '#dc2626';
    };

    const getEscalationBadge = (status: string): { bg: string; color: string } => {
      if (status === 'Level 2' || status === 'Level 3') return { bg: '#fee2e2', color: '#dc2626' };
      if (status === 'Level 1') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#f1f5f9', color: '#64748b' };
    };

    return (
      <div>
        {/* KPI Strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 28 }}>
          <div style={{ ...kpiCardStyle, borderTopColor: tc.primary }}>
            <div style={{ ...kpiValueStyle, color: tc.primary }}>{overallAckRate}%</div>
            <div style={kpiLabelStyle}>Overall Ack Rate</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#059669' }}>
            <div style={{ ...kpiValueStyle, color: '#059669' }}>{ackTarget}%</div>
            <div style={kpiLabelStyle}>Target</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#2563eb' }}>
            <div style={{ ...kpiValueStyle, color: '#2563eb' }}>{totalSent}</div>
            <div style={kpiLabelStyle}>Total Sent</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#dc2626' }}>
            <div style={{ ...kpiValueStyle, color: '#dc2626' }}>{overdueAckList.length}</div>
            <div style={kpiLabelStyle}>Overdue</div>
          </div>
        </div>

        {/* Funnel + Gauge */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Acknowledgement Funnel</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>All policies, last 90 days</span>
            </div>
            <div style={cardBodyStyle}>
              {ackFunnel.map((stage, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 16, padding: '12px 0', borderBottom: i < ackFunnel.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                  <div style={{ width: 110, fontSize: 12, fontWeight: 600, color: '#334155', flexShrink: 0 }}>{stage.stage}</div>
                  <div style={{ flex: 1, height: 28, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden' }}>
                    <div style={{ height: '100%', borderRadius: 4, width: `${stage.percent}%`, background: funnelColors[i % funnelColors.length], display: 'flex', alignItems: 'center', paddingLeft: 10, fontSize: 11, fontWeight: 700, color: '#fff' }}>
                      {stage.percent}%
                    </div>
                  </div>
                  <div style={{ minWidth: 40, textAlign: 'right', fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{stage.count.toLocaleString()}</div>
                </div>
              ))}
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Overall Acknowledgement Rate</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>vs {ackTarget}% target</span>
            </div>
            <div style={cardBodyStyle}>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
                <div style={{ width: 180, height: 180, position: 'relative' }}>
                  <svg viewBox="0 0 180 180" style={{ width: '100%', height: '100%', transform: 'rotate(-90deg)' }}>
                    <circle cx="90" cy="90" r={gaugeR} fill="none" stroke="#f1f5f9" strokeWidth="16" />
                    <circle cx="90" cy="90" r={gaugeR} fill="none" stroke={tc.primary} strokeWidth="16" strokeDasharray={`${gaugeFill} ${gaugeCirc}`} strokeLinecap="round" />
                  </svg>
                  <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
                    <div style={{ fontSize: 32, fontWeight: 700, color: tc.primary }}>{overallAckRate}%</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase' }}>Ack Rate</div>
                  </div>
                </div>
                <div style={{ fontSize: 11, color: '#64748b', marginTop: 8 }}>
                  Target: <strong style={{ color: '#d97706' }}>{ackTarget}%</strong> {' \u2014 '} {(ackTarget - overallAckRate).toFixed(1)}% below target. {Math.ceil((ackTarget - overallAckRate) / 100 * totalSent)} more acknowledgements needed.
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Department Breakdown */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Acknowledgement by Department</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>Current period</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={thStyle}>Department</th>
                  <th style={thStyle}>Total</th>
                  <th style={thStyle}>Acknowledged</th>
                  <th style={thStyle}>Pending</th>
                  <th style={thStyle}>Rate (%)</th>
                  <th style={{ ...thStyle, minWidth: 140 }}>Progress</th>
                </tr>
              </thead>
              <tbody>
                {ackByDepartment.map((dept, i) => {
                  const pending = dept.assigned - dept.acknowledged;
                  const rateColor = getRateColor(dept.rate);
                  return (
                    <tr key={i}>
                      <td style={{ ...tdStyle, fontWeight: 600 }}>{dept.department}</td>
                      <td style={tdStyle}>{dept.assigned}</td>
                      <td style={tdStyle}>{dept.acknowledged}</td>
                      <td style={tdStyle}>{pending}</td>
                      <td style={tdStyle}><span style={{ fontWeight: 700, color: rateColor }}>{dept.rate}%</span></td>
                      <td style={tdStyle}>
                        <div style={{ width: '100%', height: 6, background: '#f1f5f9', borderRadius: 3, overflow: 'hidden' }}>
                          <div style={{ height: '100%', borderRadius: 3, width: `${dept.rate}%`, background: rateColor }} />
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Overdue Items */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Overdue Acknowledgements</h3>
            <span style={{ fontSize: 11, color: '#dc2626', fontWeight: 600 }}>{overdueAckList.length} overdue</span>
          </div>
          <div style={cardBodyStyle}>
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 80px 1.5fr 120px', padding: '0 0 8px', borderBottom: '2px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8' }}>
              <div>Employee</div><div>Policy</div><div>Days Overdue</div><div>Department</div><div>Escalation</div>
            </div>
            {overdueAckList.map((item) => {
              const escBadge = getEscalationBadge(item.escalationStatus);
              return (
                <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 80px 1.5fr 120px', padding: '10px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12, alignItems: 'center' }}>
                  <div style={{ fontWeight: 600 }}>{item.userName}</div>
                  <div>{item.policyTitle}</div>
                  <div style={{ color: item.daysOverdue >= 14 ? '#dc2626' : '#d97706', fontWeight: 700 }}>{item.daysOverdue} days</div>
                  <div>{item.department}</div>
                  <div><span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: escBadge.bg, color: escBadge.color }}>{item.escalationStatus === 'None' ? 'Reminder Sent' : item.escalationStatus}</span></div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 4: SLA TRACKING
  // ============================================================================

  private _renderSLATracking(): React.ReactElement {
    const { slaMetrics, slaBreaches, slaDeptComparison } = this.state;

    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };
    const thStyle: React.CSSProperties = { fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', padding: '8px 12px', textAlign: 'left', borderBottom: '2px solid #e2e8f0' };
    const tdStyle: React.CSSProperties = { padding: '10px 12px', fontSize: 12, borderBottom: '1px solid #f1f5f9' };

    const getSlaColor = (pct: number): string => pct >= 90 ? '#059669' : pct >= 80 ? '#d97706' : '#dc2626';
    const getSlaBadge = (status: string): { bg: string; color: string } => {
      if (status === 'Met') return { bg: '#dcfce7', color: '#16a34a' };
      if (status === 'At Risk') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#fee2e2', color: '#dc2626' };
    };

    const getHeatCellStyle = (val: number): React.CSSProperties => {
      let bg = '#dcfce7'; let color = '#16a34a';
      if (val < 80) { bg = '#fee2e2'; color = '#dc2626'; }
      else if (val < 90) { bg = '#fef3c7'; color = '#d97706'; }
      return { borderRadius: 4, padding: '6px 10px', display: 'inline-block', minWidth: 50, fontSize: 11, fontWeight: 700, background: bg, color, textAlign: 'center' };
    };

    return (
      <div>
        {/* SLA Metric Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${slaMetrics.length}, 1fr)`, gap: 16, marginBottom: 28 }}>
          {slaMetrics.map((sla, i) => {
            const statusBadge = getSlaBadge(sla.status);
            const pctColor = getSlaColor(sla.percentMet);
            return (
              <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: '#334155', marginBottom: 12 }}>{sla.name}</div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
                  <span style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Target</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{sla.targetDays} days</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
                  <span style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Actual Avg</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{sla.actualAvgDays} days</span>
                </div>
                <div style={{ fontSize: 28, fontWeight: 700, lineHeight: 1.1, margin: '8px 0', color: pctColor }}>{sla.percentMet}%</div>
                <div style={{ fontSize: 10, color: '#94a3b8' }}>SLA Met</div>
                <div style={{ height: 8, borderRadius: 4, background: '#f1f5f9', overflow: 'hidden', margin: '12px 0 8px' }}>
                  <div style={{ height: '100%', borderRadius: 4, width: `${sla.percentMet}%`, background: pctColor }} />
                </div>
                <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: statusBadge.bg, color: statusBadge.color }}>{sla.status}</span>
              </div>
            );
          })}
        </div>

        {/* SLA Breaches Table */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>SLA Breaches</h3>
            <span style={{ fontSize: 11, color: '#dc2626', fontWeight: 600 }}>{slaBreaches.length} active breaches</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={thStyle}>Policy</th>
                  <th style={thStyle}>SLA Type</th>
                  <th style={thStyle}>Target</th>
                  <th style={thStyle}>Actual</th>
                  <th style={thStyle}>Exceeded By</th>
                  <th style={thStyle}>Status</th>
                </tr>
              </thead>
              <tbody>
                {slaBreaches.map((breach) => {
                  const exceeded = breach.actualDays - breach.targetDays;
                  const isCritical = exceeded > 10;
                  const excColor = isCritical ? '#dc2626' : '#d97706';
                  const statusBadge = isCritical ? { bg: '#fee2e2', color: '#dc2626' } : { bg: '#fef3c7', color: '#d97706' };
                  return (
                    <tr key={breach.id}>
                      <td style={{ ...tdStyle, fontWeight: 600 }}>{breach.policyTitle}</td>
                      <td style={tdStyle}>{breach.type}</td>
                      <td style={tdStyle}>{breach.targetDays} days</td>
                      <td style={{ ...tdStyle, color: excColor, fontWeight: 600 }}>{breach.actualDays} days</td>
                      <td style={{ ...tdStyle, color: excColor, fontWeight: 700 }}>+{exceeded} days</td>
                      <td style={tdStyle}>
                        <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: statusBadge.bg, color: statusBadge.color }}>{isCritical ? 'Critical' : 'Warning'}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Department Comparison Heatmap */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Department SLA Comparison</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>Colour-coded by compliance status</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, width: 160, textAlign: 'left' }}>Department</th>
                  <th style={{ ...thStyle, textAlign: 'center' }}>Acknowledgement</th>
                  <th style={{ ...thStyle, textAlign: 'center' }}>Approval</th>
                  <th style={{ ...thStyle, textAlign: 'center' }}>Review</th>
                </tr>
              </thead>
              <tbody>
                {slaDeptComparison.map((dept, i) => (
                  <tr key={i}>
                    <td style={{ ...tdStyle, fontSize: 12, fontWeight: 600, color: '#334155', textAlign: 'left' }}>{dept.department}</td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}><span style={getHeatCellStyle(dept.ackSla)}>{dept.ackSla}%</span></td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}><span style={getHeatCellStyle(dept.approvalSla)}>{dept.approvalSla}%</span></td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}><span style={getHeatCellStyle(dept.reviewSla)}>{dept.reviewSla}%</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 5: COMPLIANCE & RISK
  // ============================================================================

  private _renderComplianceRisk(): React.ReactElement {
    const { heatmapData, riskCards, violations } = this.state;
    const categories = ['HR', 'IT', 'Compliance', 'Safety', 'Finance'];
    const categoryLabels: Record<string, string> = { HR: 'HR', IT: 'IT Security', Compliance: 'Compliance', Safety: 'Health & Safety', Finance: 'Financial' };

    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };
    const hmThStyle: React.CSSProperties = { fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', padding: '8px 10px', textAlign: 'center', borderBottom: '2px solid #e2e8f0' };
    const hmTdStyle: React.CSSProperties = { padding: '6px 10px', textAlign: 'center', borderBottom: '1px solid #f1f5f9' };

    const getHeatCell = (val: number): { label: string; bg: string; color: string } => {
      if (val >= 90) return { label: 'Low', bg: '#dcfce7', color: '#16a34a' };
      if (val >= 80) return { label: 'Med', bg: '#fef3c7', color: '#d97706' };
      return { label: 'High', bg: '#fee2e2', color: '#dc2626' };
    };

    const getRiskLevelStyle = (level: string): { bg: string; color: string } => {
      if (level === 'high') return { bg: '#fee2e2', color: '#dc2626' };
      if (level === 'medium') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#dcfce7', color: '#16a34a' };
    };
    const getRiskScoreColor = (level: string): string => level === 'high' ? '#dc2626' : level === 'medium' ? '#d97706' : '#059669';

    const getSeverityBadge = (sev: string): { bg: string; color: string } => {
      if (sev === 'Critical') return { bg: '#fee2e2', color: '#dc2626' };
      if (sev === 'Major') return { bg: '#fef3c7', color: '#d97706' };
      if (sev === 'Minor') return { bg: '#f1f5f9', color: '#64748b' };
      return { bg: '#f1f5f9', color: '#94a3b8' };
    };

    const getStatusBadge = (status: string): { bg: string; color: string } => {
      if (status === 'Open') return { bg: '#dbeafe', color: '#2563eb' };
      if (status === 'In Progress') return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#dcfce7', color: '#16a34a' };
    };

    return (
      <div>
        {/* Risk Heatmap */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Compliance Risk Heatmap</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>Department vs Category risk assessment</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ ...hmThStyle, width: 140, textAlign: 'left' }}>Department</th>
                  {categories.map((c) => <th key={c} style={hmThStyle}>{categoryLabels[c] || c}</th>)}
                </tr>
              </thead>
              <tbody>
                {heatmapData.map((row, i) => (
                  <tr key={i}>
                    <td style={{ ...hmTdStyle, textAlign: 'left', fontSize: 12, fontWeight: 600, color: '#334155' }}>{row.department}</td>
                    {[row.hr, row.it, row.compliance, row.safety, row.finance].map((val, ci) => {
                      const cell = getHeatCell(val);
                      return (
                        <td key={ci} style={hmTdStyle}>
                          <span style={{ borderRadius: 4, padding: '6px 8px', display: 'inline-block', minWidth: 44, fontSize: 10, fontWeight: 700, background: cell.bg, color: cell.color }}>{cell.label}</span>
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Risk Category Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${Math.min(riskCards.length, 4)}, 1fr)`, gap: 16, marginBottom: 20 }}>
          {riskCards.slice(0, 4).map((risk, i) => {
            const levelStyle = getRiskLevelStyle(risk.level);
            const scoreColor = getRiskScoreColor(risk.level);
            return (
              <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 18 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: '#334155', marginBottom: 10 }}>{risk.category}</div>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginBottom: 8 }}>
                  <span style={{ fontSize: 28, fontWeight: 700, color: scoreColor }}>{risk.score}</span>
                  <span style={{ fontSize: 12, color: '#94a3b8' }}>/100</span>
                  <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: levelStyle.bg, color: levelStyle.color }}>{risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}</span>
                </div>
                <div style={{ height: 6, borderRadius: 3, background: '#f1f5f9', margin: '10px 0', overflow: 'hidden' }}>
                  <div style={{ height: '100%', borderRadius: 3, width: `${risk.score}%`, background: scoreColor }} />
                </div>
                <ul style={{ listStyle: 'none', padding: 0, margin: '8px 0 0' }}>
                  {risk.factors.map((f, fi) => (
                    <li key={fi} style={{ fontSize: 11, color: '#64748b', padding: '3px 0', paddingLeft: 14, position: 'relative' }}>
                      <span style={{ position: 'absolute', left: 0, top: 8, width: 5, height: 5, borderRadius: '50%', background: '#94a3b8' }} />
                      {f}
                    </li>
                  ))}
                </ul>
                <div style={{ fontSize: 10, color: tc.primary, fontStyle: 'italic', marginTop: 8, paddingTop: 8, borderTop: '1px solid #f1f5f9' }}>
                  Mitigation: {risk.mitigation}
                </div>
              </div>
            );
          })}
        </div>

        {/* Violations List */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Policy Violations</h3>
            <span style={{ fontSize: 11, color: '#dc2626', fontWeight: 600 }}>{violations.filter(v => v.status !== 'Resolved').length} active</span>
          </div>
          <div style={cardBodyStyle}>
            <div style={{ display: 'grid', gridTemplateColumns: '90px 3fr 1.5fr 100px 100px', padding: '0 0 8px', borderBottom: '2px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8' }}>
              <div>Severity</div><div>Policy Title</div><div>Department</div><div>Status</div><div>Detected</div>
            </div>
            {violations.map((v) => {
              const sevBadge = getSeverityBadge(v.severity);
              const statusBadge = getStatusBadge(v.status);
              return (
                <div key={v.id} style={{ display: 'grid', gridTemplateColumns: '90px 3fr 1.5fr 100px 100px', padding: '10px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12, alignItems: 'center' }}>
                  <div><span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: sevBadge.bg, color: sevBadge.color }}>{v.severity}</span></div>
                  <div style={{ fontWeight: 600 }}>{v.policyTitle}</div>
                  <div>{v.department}</div>
                  <div><span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: statusBadge.bg, color: statusBadge.color }}>{v.status}</span></div>
                  <div style={{ color: '#94a3b8' }}>{v.detectedDate}</div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 6: AUDIT & REPORTS
  // ============================================================================

  private _renderAuditReports(): React.ReactElement {
    const { auditEntries, auditFilter, scheduledReports } = this.state;

    const filteredEntries = auditFilter === 'all'
      ? auditEntries
      : auditEntries.filter(e => e.category === auditFilter);

    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };
    const thStyle: React.CSSProperties = { fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', padding: '8px 12px', textAlign: 'left', borderBottom: '2px solid #e2e8f0' };
    const tdStyle: React.CSSProperties = { padding: '10px 12px', fontSize: 12, borderBottom: '1px solid #f1f5f9', verticalAlign: 'middle' };
    const filterBtnStyle: React.CSSProperties = { padding: '8px 12px', border: '1px solid #e2e8f0', borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: 'pointer', background: '#fff', color: '#334155', fontFamily: 'inherit' };
    const filterBtnActiveStyle: React.CSSProperties = { ...filterBtnStyle, background: tc.primary, color: '#fff', borderColor: tc.primary };

    const getActionBadge = (action: string): { bg: string; color: string } => {
      const a = action.toLowerCase();
      if (a.includes('publish')) return { bg: '#dcfce7', color: '#16a34a' };
      if (a.includes('approv')) return { bg: '#dbeafe', color: '#2563eb' };
      if (a.includes('reject')) return { bg: '#fee2e2', color: '#dc2626' };
      if (a.includes('update') || a.includes('config')) return { bg: '#fef3c7', color: '#d97706' };
      if (a.includes('acknowledge') || a.includes('complete')) return { bg: tc.primaryLighter, color: tc.primary };
      if (a.includes('create') || a.includes('submit')) return { bg: '#e0e7ff', color: '#4f46e5' };
      if (a.includes('violation') || a.includes('escalat')) return { bg: '#fee2e2', color: '#dc2626' };
      if (a.includes('access') || a.includes('role')) return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#f1f5f9', color: '#64748b' };
    };

    const avatarColors = [tc.primary, tc.accent, '#7c3aed', tc.success, tc.warning, tc.danger, '#94a3b8'];
    const getInitials = (name: string): string => {
      const parts = name.split(' ');
      return parts.length >= 2 ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase() : name.substring(0, 2).toUpperCase();
    };

    const getReportCategoryBadge = (type: string): { bg: string; color: string } => {
      if (type.toLowerCase().includes('compliance')) return { bg: '#dbeafe', color: '#2563eb' };
      if (type.toLowerCase().includes('ack')) return { bg: '#dcfce7', color: '#16a34a' };
      if (type.toLowerCase().includes('audit')) return { bg: '#fef3c7', color: '#d97706' };
      return { bg: '#f1f5f9', color: '#64748b' };
    };

    const getFormatBadge = (fmt: string): { bg: string; color: string } => {
      if (fmt === 'PDF') return { bg: '#fee2e2', color: '#dc2626' };
      if (fmt === 'Excel') return { bg: '#dcfce7', color: '#16a34a' };
      return { bg: '#f1f5f9', color: '#64748b' };
    };

    return (
      <div>
        {/* Filters */}
        <div style={{ display: 'flex', gap: 12, marginBottom: 20, alignItems: 'center' }}>
          {['all', 'policy', 'user', 'system', 'compliance', 'access'].map((cat) => (
            <button
              key={cat}
              style={auditFilter === cat ? filterBtnActiveStyle : filterBtnStyle}
              onClick={() => this.setState({ auditFilter: cat })}
            >
              {cat === 'all' ? 'All Categories' : cat.charAt(0).toUpperCase() + cat.slice(1)}
            </button>
          ))}
          <div style={{ flex: 1 }} />
          <button style={filterBtnActiveStyle} onClick={() => {
            const rows = (this.state as any).auditEntries || [];
            if (rows.length === 0) return;
            const headers = ['Timestamp', 'User', 'Action', 'Category', 'Resource', 'Department'];
            const csv = [headers.join(','), ...rows.map((r: any) => [r.timestamp || '', r.user || '', r.action || '', r.category || '', r.resource || '', r.department || ''].map((v: string) => `"${v.replace(/"/g, '""')}"`).join(','))].join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a'); link.href = url; link.download = `AuditLog_${new Date().toISOString().split('T')[0]}.csv`; link.click(); URL.revokeObjectURL(url);
          }}>Export</button>
        </div>

        {/* Audit Log Table */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Audit Log</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>{filteredEntries.length} events</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={thStyle}>Timestamp</th>
                  <th style={thStyle}>User</th>
                  <th style={thStyle}>Action</th>
                  <th style={thStyle}>Category</th>
                  <th style={thStyle}>Resource</th>
                  <th style={thStyle}>Department</th>
                </tr>
              </thead>
              <tbody>
                {filteredEntries.map((entry, idx) => {
                  const actionBadge = getActionBadge(entry.action);
                  return (
                    <tr key={entry.id}>
                      <td style={{ ...tdStyle, color: '#94a3b8', whiteSpace: 'nowrap' }}>{entry.timestamp}</td>
                      <td style={tdStyle}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <div style={{ width: 28, height: 28, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700, color: '#fff', flexShrink: 0, background: avatarColors[idx % avatarColors.length] }}>
                            {getInitials(entry.userName)}
                          </div>
                          <span>{entry.userName}</span>
                        </div>
                      </td>
                      <td style={tdStyle}>
                        <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: actionBadge.bg, color: actionBadge.color }}>{entry.action}</span>
                      </td>
                      <td style={tdStyle}>{entry.category.charAt(0).toUpperCase() + entry.category.slice(1)}</td>
                      <td style={{ ...tdStyle, fontWeight: 600 }}>{entry.resourceTitle}</td>
                      <td style={tdStyle}>{entry.department}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Scheduled Reports */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Scheduled Reports</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>{scheduledReports.length} active schedules</span>
          </div>
          <div style={cardBodyStyle}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 20 }}>
              {scheduledReports.slice(0, 3).map((rpt) => {
                const catBadge = getReportCategoryBadge(rpt.type);
                const fmtBadge = getFormatBadge(rpt.format);
                return (
                  <div key={rpt.id} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: '#0f172a' }}>{rpt.title}</div>
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: catBadge.bg, color: catBadge.color }}>{rpt.type}</span>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0' }}>
                      <span style={{ fontSize: 11, color: '#94a3b8' }}>Schedule</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>{rpt.schedule}</span>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0' }}>
                      <span style={{ fontSize: 11, color: '#94a3b8' }}>Last Run</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>{rpt.lastRun}</span>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0' }}>
                      <span style={{ fontSize: 11, color: '#94a3b8' }}>Next Run</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>{rpt.nextRun}</span>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0' }}>
                      <span style={{ fontSize: 11, color: '#94a3b8' }}>Format</span>
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: fmtBadge.bg, color: fmtBadge.color }}>{rpt.format}</span>
                    </div>
                    <div style={{ marginTop: 12, paddingTop: 12, borderTop: '1px solid #f1f5f9', display: 'flex', gap: 8 }}>
                      <button style={{ ...filterBtnStyle, fontSize: 11, padding: '4px 10px' }} onClick={() => { window.location.href = '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=reports'; }}>Run Now</button>
                      <button style={{ ...filterBtnStyle, fontSize: 11, padding: '4px 10px' }} onClick={() => { window.location.href = '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=reports'; }}>Edit</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAB 7: QUIZ ANALYTICS
  // ============================================================================

  private _renderQuizAnalytics(): React.ReactElement {
    const { quizOverview, quizPerformance, quizByDepartment, quizTopPerformers } = this.state;

    const kpiCardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 16px', borderTop: '3px solid transparent' };
    const kpiValueStyle: React.CSSProperties = { fontSize: 28, fontWeight: 700, lineHeight: 1.1 };
    const kpiLabelStyle: React.CSSProperties = { fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 };
    const kpiTrendStyle: React.CSSProperties = { fontSize: 10, marginTop: 6, display: 'flex', alignItems: 'center', gap: 4 };
    const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 };
    const cardHeaderStyle: React.CSSProperties = { padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' };
    const cardBodyStyle: React.CSSProperties = { padding: 20 };
    const h3Style: React.CSSProperties = { fontSize: 14, fontWeight: 700, margin: 0 };
    const thStyle: React.CSSProperties = { fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', padding: '8px 12px', textAlign: 'left', borderBottom: '2px solid #e2e8f0' };
    const tdStyle: React.CSSProperties = { padding: '10px 12px', fontSize: 12, borderBottom: '1px solid #f1f5f9' };

    const difficultyBadge = (d: string): { bg: string; color: string } => {
      if (d === 'Easy') return { bg: '#dcfce7', color: '#16a34a' };
      if (d === 'Medium') return { bg: '#fef3c7', color: '#d97706' };
      if (d === 'Hard') return { bg: '#fee2e2', color: '#dc2626' };
      if (d === 'Expert') return { bg: '#ede9fe', color: '#7c3aed' };
      return { bg: '#f1f5f9', color: '#64748b' };
    };

    const getScoreColor = (score: number): string => score >= 80 ? '#059669' : score >= 70 ? '#d97706' : '#dc2626';

    const getDeptBarGradient = (score: number): string => {
      if (score >= 80) return `linear-gradient(90deg, ${tc.primary}, #14b8a6)`;
      if (score >= 70) return 'linear-gradient(90deg, #d97706, #f59e0b)';
      return 'linear-gradient(90deg, #dc2626, #ef4444)';
    };

    // Sort departments by avg score desc
    const sortedDepts = [...quizByDepartment].sort((a, b) => b.avgScore - a.avgScore);

    return (
      <div>
        {/* KPI Strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 12, marginBottom: 28 }}>
          <div style={{ ...kpiCardStyle, borderTopColor: '#2563eb' }}>
            <div style={{ ...kpiValueStyle, color: '#2563eb' }}>{quizOverview.totalQuizzes}</div>
            <div style={kpiLabelStyle}>Total Quizzes</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#059669' }}>
            <div style={{ ...kpiValueStyle, color: '#059669' }}>{quizOverview.activeQuizzes}</div>
            <div style={kpiLabelStyle}>Active</div>
            <div style={{ ...kpiTrendStyle, color: '#94a3b8' }}>{quizOverview.totalQuizzes - quizOverview.activeQuizzes} draft</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: tc.primary }}>
            <div style={{ ...kpiValueStyle, color: tc.primary }}>{quizOverview.totalAttempts.toLocaleString()}</div>
            <div style={kpiLabelStyle}>Total Attempts</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#d97706' }}>
            <div style={{ ...kpiValueStyle, color: '#d97706' }}>{quizOverview.avgScore}%</div>
            <div style={kpiLabelStyle}>Avg Score</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#7c3aed' }}>
            <div style={{ ...kpiValueStyle, color: '#7c3aed' }}>{quizOverview.passRate}%</div>
            <div style={kpiLabelStyle}>Pass Rate</div>
          </div>
          <div style={{ ...kpiCardStyle, borderTopColor: '#94a3b8' }}>
            <div style={{ ...kpiValueStyle, color: '#94a3b8' }}>{quizOverview.avgCompletionTime}</div>
            <div style={kpiLabelStyle}>Avg Time</div>
          </div>
        </div>

        {/* Quiz Performance Table */}
        <div style={cardStyle}>
          <div style={cardHeaderStyle}>
            <h3 style={h3Style}>Quiz Performance</h3>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>All active quizzes</span>
          </div>
          <div style={cardBodyStyle}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={thStyle}>Quiz Title</th>
                  <th style={thStyle}>Attempts</th>
                  <th style={thStyle}>Avg Score</th>
                  <th style={thStyle}>Pass Rate</th>
                  <th style={thStyle}>Avg Time</th>
                  <th style={thStyle}>Difficulty</th>
                </tr>
              </thead>
              <tbody>
                {quizPerformance.map((quiz, i) => {
                  const dBadge = difficultyBadge(quiz.difficulty);
                  return (
                    <tr key={i}>
                      <td style={{ ...tdStyle, fontWeight: 600 }}>{quiz.title}</td>
                      <td style={tdStyle}>{quiz.attempts}</td>
                      <td style={{ ...tdStyle, color: getScoreColor(quiz.avgScore), fontWeight: 700 }}>{quiz.avgScore}%</td>
                      <td style={tdStyle}>{quiz.passRate}%</td>
                      <td style={tdStyle}>{quiz.avgTime}</td>
                      <td style={tdStyle}>
                        <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: dBadge.bg, color: dBadge.color }}>{quiz.difficulty}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Dept Scores + Top Performers */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Average Score by Department</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>All quizzes combined</span>
            </div>
            <div style={cardBodyStyle}>
              {sortedDepts.map((dept, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '8px 0', borderBottom: '1px solid #f8fafc' }}>
                  <div style={{ width: 120, fontSize: 12, fontWeight: 500, color: '#334155', flexShrink: 0 }}>{dept.department}</div>
                  <div style={{ flex: 1, height: 24, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden' }}>
                    <div style={{ height: '100%', borderRadius: 4, width: `${dept.avgScore}%`, background: getDeptBarGradient(dept.avgScore), display: 'flex', alignItems: 'center', paddingLeft: 8, fontSize: 10, fontWeight: 700, color: '#fff' }}>
                      {dept.avgScore}%
                    </div>
                  </div>
                  <div style={{ minWidth: 40, textAlign: 'right', fontSize: 12, fontWeight: 700, color: '#0f172a' }}>{dept.avgScore}%</div>
                </div>
              ))}
            </div>
          </div>

          <div style={cardStyle}>
            <div style={cardHeaderStyle}>
              <h3 style={h3Style}>Top Performers</h3>
              <span style={{ fontSize: 11, color: '#94a3b8' }}>Highest average scores</span>
            </div>
            <div style={cardBodyStyle}>
              {quizTopPerformers.map((p, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 0', borderBottom: i < quizTopPerformers.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                  <div style={{ width: 28, height: 28, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0, background: i === 0 ? '#fef3c7' : '#f1f5f9', color: i === 0 ? '#d97706' : '#64748b' }}>{i + 1}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{p.name}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{p.department}</div>
                  </div>
                  <div style={{ display: 'flex', gap: 16 }}>
                    <div style={{ textAlign: 'right' }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: '#059669' }}>{p.avgScore}%</div>
                      <div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase' }}>Avg Score</div>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <div style={{ fontSize: 13, fontWeight: 700 }}>{p.quizzesCompleted}</div>
                      <div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase' }}>Completed</div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
