// @ts-nocheck
import * as React from 'react';
import styles from './PolicyAnalytics.module.scss';
import { IPolicyAnalyticsProps } from './IPolicyAnalyticsProps';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import {
  Pivot,
  PivotItem,
  Icon,
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

  public render(): React.ReactElement<IPolicyAnalyticsProps> {
    const { activeTab } = this.state;

    return (
      <JmlAppLayout
        title={this.props.title}
        context={this.props.context}
        sp={this.props.sp}
        pageTitle="Policy Analytics"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Analytics' }]}
      >
        <div className={styles.policyAnalytics}>
          {/* Tab Navigation */}
          <div className={styles.tabSection}>
            <Pivot
              selectedKey={activeTab}
              onLinkClick={(item) => { if (item) this.setState({ activeTab: item.props.itemKey || 'executive' }); }}
              styles={{
                root: { borderBottom: '1px solid #e2e8f0', paddingLeft: 40 },
                link: { fontSize: 13, fontWeight: 400, color: '#64748b', height: 44 },
                linkIsSelected: { fontSize: 13, fontWeight: 400, color: '#0d9488' },
              }}
            >
              <PivotItem headerText="Executive Dashboard" itemKey="executive" itemIcon="ViewDashboard" />
              <PivotItem headerText="Policy Metrics" itemKey="metrics" itemIcon="BarChartVertical" />
              <PivotItem headerText="Acknowledgement Tracking" itemKey="acknowledgements" itemIcon="CheckMark" />
              <PivotItem headerText="SLA Tracking" itemKey="sla" itemIcon="Timer" />
              <PivotItem headerText="Compliance & Risk" itemKey="compliance" itemIcon="Shield" />
              <PivotItem headerText="Audit & Reports" itemKey="audit" itemIcon="ReportDocument" />
              <PivotItem headerText="Quiz Analytics" itemKey="quiz" itemIcon="Questionnaire" />
            </Pivot>
          </div>

          {/* Tab Content */}
          <div className={styles.tabContent}>
            {activeTab === 'executive' && this._renderExecutiveDashboard()}
            {activeTab === 'metrics' && this._renderPolicyMetrics()}
            {activeTab === 'acknowledgements' && this._renderAcknowledgementTracking()}
            {activeTab === 'sla' && this._renderSLATracking()}
            {activeTab === 'compliance' && this._renderComplianceRisk()}
            {activeTab === 'audit' && this._renderAuditReports()}
            {activeTab === 'quiz' && this._renderQuizAnalytics()}
          </div>
        </div>
      </JmlAppLayout>
    );
  }

  // ============================================================================
  // TAB 1: EXECUTIVE DASHBOARD
  // ============================================================================

  private _renderExecutiveDashboard(): React.ReactElement {
    const { overallCompliance, activePolicies, pendingReviews, overdueAcks, criticalViolations, avgResolutionDays, complianceTrend, riskIndicators, alerts, deadlines } = this.state;

    return (
      <div className={styles.executiveTab}>
        {/* KPI Cards */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}>
              <h3>Key Performance Indicators</h3>
            </div>
            <div className={styles.kpiGrid}>
              <div className={`${styles.kpiCard} ${styles.kpiPrimary}`}>
                <div className={styles.kpiValue}>{overallCompliance}%</div>
                <div className={styles.kpiLabel}>Overall Compliance</div>
                <div className={styles.kpiTrend}>
                  <Icon iconName="CaretSolidUp" className={styles.trendUp} />
                  <span className={styles.trendUp}>+2.1% vs last month</span>
                </div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{activePolicies}</div>
                <div className={styles.kpiLabel}>Active Policies</div>
              </div>
              <div className={`${styles.kpiCard} ${styles.kpiWarning}`}>
                <div className={styles.kpiValue}>{pendingReviews}</div>
                <div className={styles.kpiLabel}>Pending Reviews</div>
              </div>
              <div className={`${styles.kpiCard} ${styles.kpiDanger}`}>
                <div className={styles.kpiValue}>{overdueAcks}</div>
                <div className={styles.kpiLabel}>Overdue Acknowledgements</div>
              </div>
              <div className={`${styles.kpiCard} ${styles.kpiDanger}`}>
                <div className={styles.kpiValue}>{criticalViolations}</div>
                <div className={styles.kpiLabel}>Critical Violations</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{avgResolutionDays}</div>
                <div className={styles.kpiLabel}>Avg Resolution (Days)</div>
              </div>
            </div>
          </div>
        </div>

        {/* Compliance Trend */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}>
              <h3>Compliance Trend <small>(12 Months)</small></h3>
            </div>
            <div className={styles.trendChart}>
              <div className={styles.trendBars}>
                {complianceTrend.map((pt, idx) => (
                  <div key={idx} className={styles.trendBarCol}>
                    <div className={styles.trendBarValue}>{pt.value}%</div>
                    <div className={styles.trendBar} style={{ height: `${(pt.value / 100) * 160}px` }} />
                    <div className={styles.trendBarLabel}>{pt.month}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Risk Indicators & Alerts */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.twoColumnSection}>
              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Risk Indicators</h3></div>
                <div className={styles.riskGrid}>
                  {riskIndicators.map((risk, idx) => (
                    <div key={idx} className={`${styles.riskCard} ${styles[`risk${risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}`]}`}>
                      <div className={styles.riskHeader}>
                        <span className={styles.riskCategory}>{risk.category}</span>
                        <span className={`${styles.riskBadge} ${styles[`badge${risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}`]}`}>
                          {risk.level.toUpperCase()}
                        </span>
                      </div>
                      <div className={styles.riskScore}>
                        <div className={styles.riskScoreBar}>
                          <div className={styles.riskScoreFill} style={{ width: `${risk.score}%` }} />
                        </div>
                        <span>{risk.score}/100</span>
                      </div>
                      <div className={styles.riskTrend}>
                        <Icon iconName={risk.trend === 'improving' ? 'CaretSolidDown' : risk.trend === 'worsening' ? 'CaretSolidUp' : 'Remove'} />
                        <span>{risk.trend.charAt(0).toUpperCase() + risk.trend.slice(1)}</span>
                      </div>
                      <div className={styles.riskMitigation}>{risk.mitigation}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Alerts &amp; Upcoming Deadlines</h3></div>
                <div className={styles.alertList}>
                  {alerts.map((alert) => (
                    <div key={alert.id} className={`${styles.alertItem} ${styles[`alert${alert.type.charAt(0).toUpperCase() + alert.type.slice(1)}`]}`}>
                      <Icon iconName={alert.type === 'critical' ? 'ErrorBadge' : alert.type === 'warning' ? 'Warning' : 'Info'} className={styles.alertIcon} />
                      <div className={styles.alertContent}>
                        <div className={styles.alertTitle}>{alert.title}</div>
                        <div className={styles.alertMessage}>{alert.message}</div>
                        <div className={styles.alertDate}>{alert.date}</div>
                      </div>
                    </div>
                  ))}
                </div>
                <table className={styles.dataTable} style={{ marginTop: 16 }}>
                  <thead>
                    <tr>
                      <th>Item</th>
                      <th>Due Date</th>
                      <th>Days Left</th>
                      <th>Priority</th>
                    </tr>
                  </thead>
                  <tbody>
                    {deadlines.map((dl) => (
                      <tr key={dl.id}>
                        <td className={styles.cellTitle}>{dl.title}</td>
                        <td>{dl.dueDate}</td>
                        <td className={dl.daysRemaining <= 7 ? styles.cellDanger : ''}>{dl.daysRemaining}</td>
                        <td><span className={`${styles.priorityBadge} ${styles[`priority${dl.priority.charAt(0).toUpperCase() + dl.priority.slice(1)}`]}`}>{dl.priority}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
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

    return (
      <div className={styles.metricsTab}>
        {/* Policy by Status */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Policy Status Breakdown</h3></div>
            <div className={styles.statusRow}>
              {policyByStatus.map((s, i) => (
                <div key={i} className={styles.statusCard}>
                  <div className={styles.statusDot} style={{ background: s.color }} />
                  <div className={styles.statusInfo}>
                    <div className={styles.statusCount}>{s.count}</div>
                    <div className={styles.statusLabel}>{s.status}</div>
                  </div>
                  <div className={styles.statusPercent}>{((s.count / totalPolicies) * 100).toFixed(0)}%</div>
                </div>
              ))}
            </div>
            <div className={styles.statusBarContainer}>
              {policyByStatus.map((s, i) => (
                <div
                  key={i}
                  className={styles.statusBarSegment}
                  style={{ width: `${(s.count / totalPolicies) * 100}%`, background: s.color }}
                  title={`${s.status}: ${s.count}`}
                />
              ))}
            </div>
          </div>
        </div>

        {/* Policy by Category */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Policies by Category</h3></div>
            <div className={styles.categoryBars}>
              {policyByCategory.map((cat, i) => {
                const maxCount = policyByCategory[0].count;
                return (
                  <div key={i} className={styles.categoryRow}>
                    <div className={styles.categoryName}>{cat.category}</div>
                    <div className={styles.categoryBarOuter}>
                      <div className={styles.categoryBarInner} style={{ width: `${(cat.count / maxCount) * 100}%` }} />
                    </div>
                    <div className={styles.categoryCount}>{cat.count}</div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* Most Viewed & Recently Published */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.twoColumnSection}>
              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Most Viewed Policies</h3></div>
                <table className={styles.dataTable}>
                  <thead>
                    <tr><th>#</th><th>Policy</th><th>Category</th><th>Views</th></tr>
                  </thead>
                  <tbody>
                    {mostViewed.map((p, i) => (
                      <tr key={i}>
                        <td className={styles.cellRank}>{i + 1}</td>
                        <td className={styles.cellTitle}>{p.title}</td>
                        <td>{p.category}</td>
                        <td className={styles.cellNumber}>{p.views.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Recently Published</h3></div>
                <div className={styles.timelineList}>
                  {recentlyPublished.map((p, i) => (
                    <div key={i} className={styles.timelineItem}>
                      <div className={styles.timelineDot} />
                      <div className={styles.timelineContent}>
                        <div className={styles.timelineTitle}>{p.title}</div>
                        <div className={styles.timelineMeta}>{p.date} by {p.author}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Policy Aging */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Policy Aging</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr><th>Age Range</th><th>Count</th><th>Overdue</th></tr>
              </thead>
              <tbody>
                {policyAging.map((a, i) => (
                  <tr key={i}>
                    <td>{a.range}</td>
                    <td className={styles.cellNumber}>{a.count}</td>
                    <td className={a.overdue > 0 ? styles.cellDanger : ''}>{a.overdue}</td>
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
  // TAB 3: ACKNOWLEDGEMENT TRACKING
  // ============================================================================

  private _renderAcknowledgementTracking(): React.ReactElement {
    const { overallAckRate, ackTarget, ackFunnel, ackByDepartment, overdueAckList } = this.state;
    const ackGap = ackTarget - overallAckRate;

    return (
      <div className={styles.ackTab}>
        {/* Big Rate Indicator */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.ackHero}>
              <div className={styles.ackRateCircle}>
                <div className={styles.ackRateValue}>{overallAckRate}%</div>
                <div className={styles.ackRateLabel}>Overall Acknowledgement Rate</div>
              </div>
              <div className={styles.ackTargetInfo}>
                <div className={styles.ackTargetLine}>
                  <span>Target SLA:</span> <span>{ackTarget}%</span>
                </div>
                <div className={styles.ackTargetLine}>
                  <span>Gap to Target:</span>
                  <span className={styles.cellDanger}>{ackGap.toFixed(1)}%</span>
                </div>
                <div className={styles.ackTargetLine}>
                  <span>Status:</span>
                  <span className={`${styles.slaBadge} ${styles.slaAtRisk}`}>At Risk</span>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Acknowledgement Funnel */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Acknowledgement Funnel</h3></div>
            <div className={styles.funnelContainer}>
              {ackFunnel.map((stage, i) => (
                <div key={i} className={styles.funnelStep}>
                  <div className={styles.funnelBar} style={{ width: `${stage.percent}%` }}>
                    <span className={styles.funnelLabel}>{stage.stage}</span>
                    <span className={styles.funnelValue}>{stage.count.toLocaleString()} ({stage.percent}%)</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Department Scorecard */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Department Scorecard</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Department</th>
                  <th>Assigned</th>
                  <th>Acknowledged</th>
                  <th>Rate</th>
                  <th>SLA Status</th>
                </tr>
              </thead>
              <tbody>
                {ackByDepartment.map((dept, i) => (
                  <tr key={i}>
                    <td className={styles.cellTitle}>{dept.department}</td>
                    <td className={styles.cellNumber}>{dept.assigned}</td>
                    <td className={styles.cellNumber}>{dept.acknowledged}</td>
                    <td className={styles.cellNumber}>
                      <div className={styles.rateBar}>
                        <div className={styles.rateBarFill} style={{ width: `${dept.rate}%`, background: dept.rate >= 95 ? '#10b981' : dept.rate >= 90 ? '#f59e0b' : '#ef4444' }} />
                      </div>
                      {dept.rate}%
                    </td>
                    <td>
                      <span className={`${styles.slaBadge} ${styles[`sla${dept.slaStatus.replace(' ', '')}`]}`}>
                        {dept.slaStatus}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Overdue Acknowledgements */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Overdue Acknowledgements</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>User</th>
                  <th>Policy</th>
                  <th>Days Overdue</th>
                  <th>Department</th>
                  <th>Escalation</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {overdueAckList.map((item) => (
                  <tr key={item.id}>
                    <td className={styles.cellTitle}>{item.userName}</td>
                    <td>{item.policyTitle}</td>
                    <td className={styles.cellDanger}>{item.daysOverdue}</td>
                    <td>{item.department}</td>
                    <td>
                      <span className={`${styles.escalationBadge} ${item.escalationStatus !== 'None' ? styles.escalationActive : ''}`}>
                        {item.escalationStatus}
                      </span>
                    </td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <IconButton iconProps={{ iconName: 'TeamsLogo' }} title={`Nudge ${item.userName} on Teams`} ariaLabel={`Nudge on Teams`} styles={{ root: { width: 28, height: 28, color: '#6264a7' }, rootHovered: { color: '#4b4d8f', background: '#f3f2f1' } }} onClick={() => alert(`Teams nudge sent to ${item.userName}`)} />
                        <IconButton iconProps={{ iconName: 'Mail' }} title={`Email ${item.userName}`} ariaLabel={`Send email reminder`} styles={{ root: { width: 28, height: 28, color: '#0078d4' }, rootHovered: { color: '#005a9e', background: '#f3f2f1' } }} onClick={() => alert(`Email reminder sent to ${item.userName}`)} />
                      </div>
                    </td>
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
  // TAB 4: SLA TRACKING
  // ============================================================================

  private _renderSLATracking(): React.ReactElement {
    const { slaMetrics, slaBreaches, slaDeptComparison } = this.state;

    return (
      <div className={styles.slaTab}>
        {/* SLA Summary Cards */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>SLA Summary</h3></div>
            <div className={styles.slaCardGrid}>
              {slaMetrics.map((sla, i) => (
                <div key={i} className={`${styles.slaCard} ${styles[`slaCard${sla.status.replace(' ', '')}`]}`}>
                  <div className={styles.slaCardHeader}>
                    <span className={styles.slaCardName}>{sla.name}</span>
                    <span className={`${styles.slaBadge} ${styles[`sla${sla.status.replace(' ', '')}`]}`}>{sla.status}</span>
                  </div>
                  <div className={styles.slaCardBody}>
                    <div className={styles.slaMetricRow}>
                      <span>Target</span><span>{sla.targetDays} days</span>
                    </div>
                    <div className={styles.slaMetricRow}>
                      <span>Actual Avg</span><span>{sla.actualAvgDays} days</span>
                    </div>
                    <div className={styles.slaMetricRow}>
                      <span>% Met</span>
                      <span className={sla.percentMet >= 90 ? styles.textSuccess : sla.percentMet >= 80 ? styles.textWarning : styles.textDanger}>{sla.percentMet}%</span>
                    </div>
                  </div>
                  <div className={styles.slaCardBar}>
                    <div className={styles.slaCardBarFill} style={{ width: `${sla.percentMet}%`, background: sla.percentMet >= 90 ? '#10b981' : sla.percentMet >= 80 ? '#f59e0b' : '#ef4444' }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* SLA Breach Log */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>SLA Breach Log</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Policy</th>
                  <th>Type</th>
                  <th>Target (Days)</th>
                  <th>Actual (Days)</th>
                  <th>Breached Date</th>
                  <th>Department</th>
                </tr>
              </thead>
              <tbody>
                {slaBreaches.map((breach) => (
                  <tr key={breach.id}>
                    <td className={styles.cellTitle}>{breach.policyTitle}</td>
                    <td>{breach.type}</td>
                    <td className={styles.cellNumber}>{breach.targetDays}</td>
                    <td className={styles.cellDanger}>{breach.actualDays}</td>
                    <td>{breach.breachedDate}</td>
                    <td>{breach.department}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Department SLA Comparison */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Department SLA Comparison</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Department</th>
                  <th>Review SLA %</th>
                  <th>Acknowledgement SLA %</th>
                  <th>Approval SLA %</th>
                </tr>
              </thead>
              <tbody>
                {slaDeptComparison.map((dept, i) => (
                  <tr key={i}>
                    <td className={styles.cellTitle}>{dept.department}</td>
                    <td className={styles.cellNumber}>
                      <span className={dept.reviewSla >= 90 ? styles.textSuccess : dept.reviewSla >= 80 ? styles.textWarning : styles.textDanger}>{dept.reviewSla}%</span>
                    </td>
                    <td className={styles.cellNumber}>
                      <span className={dept.ackSla >= 90 ? styles.textSuccess : dept.ackSla >= 80 ? styles.textWarning : styles.textDanger}>{dept.ackSla}%</span>
                    </td>
                    <td className={styles.cellNumber}>
                      <span className={dept.approvalSla >= 90 ? styles.textSuccess : dept.approvalSla >= 80 ? styles.textWarning : styles.textDanger}>{dept.approvalSla}%</span>
                    </td>
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

    const getHeatColor = (val: number): string => {
      if (val >= 95) return '#059669';
      if (val >= 90) return '#10b981';
      if (val >= 85) return '#34d399';
      if (val >= 80) return '#f59e0b';
      if (val >= 75) return '#f97316';
      return '#ef4444';
    };

    return (
      <div className={styles.complianceTab}>
        {/* Compliance Heatmap */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Compliance Heatmap <small>Departments vs Categories</small></h3></div>
            <div className={styles.heatmapContainer}>
              <table className={styles.heatmapTable}>
                <thead>
                  <tr>
                    <th>Department</th>
                    {categories.map((c) => <th key={c}>{c}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {heatmapData.map((row, i) => (
                    <tr key={i}>
                      <td className={styles.cellTitle}>{row.department}</td>
                      <td style={{ background: getHeatColor(row.hr), color: '#fff', textAlign: 'center', fontWeight: 400 }}>{row.hr}%</td>
                      <td style={{ background: getHeatColor(row.it), color: '#fff', textAlign: 'center', fontWeight: 400 }}>{row.it}%</td>
                      <td style={{ background: getHeatColor(row.compliance), color: '#fff', textAlign: 'center', fontWeight: 400 }}>{row.compliance}%</td>
                      <td style={{ background: getHeatColor(row.safety), color: '#fff', textAlign: 'center', fontWeight: 400 }}>{row.safety}%</td>
                      <td style={{ background: getHeatColor(row.finance), color: '#fff', textAlign: 'center', fontWeight: 400 }}>{row.finance}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className={styles.heatmapLegend}>
                <span className={styles.legendLabel}>Legend:</span>
                <span className={styles.legendItem}><span style={{ background: '#ef4444' }} className={styles.legendDot} /> &lt;75%</span>
                <span className={styles.legendItem}><span style={{ background: '#f97316' }} className={styles.legendDot} /> 75–79%</span>
                <span className={styles.legendItem}><span style={{ background: '#f59e0b' }} className={styles.legendDot} /> 80–84%</span>
                <span className={styles.legendItem}><span style={{ background: '#34d399' }} className={styles.legendDot} /> 85–89%</span>
                <span className={styles.legendItem}><span style={{ background: '#10b981' }} className={styles.legendDot} /> 90–94%</span>
                <span className={styles.legendItem}><span style={{ background: '#059669' }} className={styles.legendDot} /> 95%+</span>
              </div>
            </div>
          </div>
        </div>

        {/* Risk Cards */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Risk Assessment</h3></div>
            <div className={styles.riskGrid}>
              {riskCards.map((risk, i) => (
                <div key={i} className={`${styles.riskCard} ${styles[`risk${risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}`]}`}>
                  <div className={styles.riskHeader}>
                    <span className={styles.riskCategory}>{risk.category}</span>
                    <span className={`${styles.riskBadge} ${styles[`badge${risk.level.charAt(0).toUpperCase() + risk.level.slice(1)}`]}`}>
                      {risk.level.toUpperCase()}
                    </span>
                  </div>
                  <div className={styles.riskScore}>
                    <div className={styles.riskScoreBar}>
                      <div className={styles.riskScoreFill} style={{ width: `${risk.score}%` }} />
                    </div>
                    <span>{risk.score}/100</span>
                  </div>
                  <div className={styles.riskFactors}>
                    {risk.factors.map((f, fi) => (
                      <div key={fi} className={styles.riskFactor}><Icon iconName="StatusCircleInner" style={{ fontSize: 6, marginRight: 6 }} />{f}</div>
                    ))}
                  </div>
                  <div className={styles.riskMitigation}>
                    <span>Mitigation:</span> {risk.mitigation}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Violation Log */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Violation Log</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Severity</th>
                  <th>Policy</th>
                  <th>Department</th>
                  <th>Status</th>
                  <th>Detected</th>
                </tr>
              </thead>
              <tbody>
                {violations.map((v) => (
                  <tr key={v.id}>
                    <td>
                      <span className={`${styles.severityBadge} ${styles[`severity${v.severity}`]}`}>
                        {v.severity}
                      </span>
                    </td>
                    <td className={styles.cellTitle}>{v.policyTitle}</td>
                    <td>{v.department}</td>
                    <td>
                      <span className={`${styles.statusBadge} ${styles[`status${v.status.replace(' ', '')}`]}`}>
                        {v.status}
                      </span>
                    </td>
                    <td>{v.detectedDate}</td>
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
  // TAB 6: AUDIT & REPORTS
  // ============================================================================

  private _renderAuditReports(): React.ReactElement {
    const { auditEntries, auditFilter, scheduledReports } = this.state;

    const filteredEntries = auditFilter === 'all'
      ? auditEntries
      : auditEntries.filter(e => e.category === auditFilter);

    return (
      <div className={styles.auditTab}>
        {/* Audit Trail */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}>
              <h3>Audit Trail</h3>
              <div className={styles.auditFilters}>
                {['all', 'policy', 'user', 'system', 'compliance', 'access'].map((cat) => (
                  <button
                    key={cat}
                    className={`${styles.filterBtn} ${auditFilter === cat ? styles.filterBtnActive : ''}`}
                    onClick={() => this.setState({ auditFilter: cat })}
                  >
                    {cat.charAt(0).toUpperCase() + cat.slice(1)}
                  </button>
                ))}
              </div>
            </div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Timestamp</th>
                  <th>User</th>
                  <th>Action</th>
                  <th>Category</th>
                  <th>Resource</th>
                  <th>Department</th>
                </tr>
              </thead>
              <tbody>
                {filteredEntries.map((entry) => (
                  <tr key={entry.id}>
                    <td className={styles.cellMono}>{entry.timestamp}</td>
                    <td className={styles.cellTitle}>{entry.userName}</td>
                    <td>{entry.action}</td>
                    <td>
                      <span className={`${styles.categoryBadge} ${styles[`cat${entry.category.charAt(0).toUpperCase() + entry.category.slice(1)}`]}`}>
                        {entry.category}
                      </span>
                    </td>
                    <td>{entry.resourceTitle}</td>
                    <td>{entry.department}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Generate Reports */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Generate Reports</h3></div>
            <div className={styles.reportButtons}>
              <PrimaryButton iconProps={{ iconName: 'ReportDocument' }} text="Compliance Report" onClick={() => alert('Generating Compliance Report...')} />
              <DefaultButton iconProps={{ iconName: 'DownloadDocument' }} text="Export Audit Log" onClick={() => alert('Exporting Audit Log...')} />
              <DefaultButton iconProps={{ iconName: 'People' }} text="Department Report" onClick={() => alert('Generating Department Report...')} />
              <DefaultButton iconProps={{ iconName: 'BarChartVertical' }} text="Trend Analysis" onClick={() => alert('Generating Trend Analysis...')} />
            </div>
          </div>
        </div>

        {/* Scheduled Reports */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Scheduled Reports</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Report Name</th>
                  <th>Type</th>
                  <th>Schedule</th>
                  <th>Last Run</th>
                  <th>Next Run</th>
                  <th>Format</th>
                </tr>
              </thead>
              <tbody>
                {scheduledReports.map((rpt) => (
                  <tr key={rpt.id}>
                    <td className={styles.cellTitle}>{rpt.title}</td>
                    <td>{rpt.type}</td>
                    <td>{rpt.schedule}</td>
                    <td>{rpt.lastRun}</td>
                    <td>{rpt.nextRun}</td>
                    <td><span className={styles.formatBadge}>{rpt.format}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Export Options */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Export Options</h3></div>
            <div className={styles.exportRow}>
              <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export to Excel" onClick={() => alert('Exporting to Excel...')} />
              <DefaultButton iconProps={{ iconName: 'PDF' }} text="Export to PDF" onClick={() => alert('Exporting to PDF...')} />
              <DefaultButton iconProps={{ iconName: 'TextDocument' }} text="Export to CSV" onClick={() => alert('Exporting to CSV...')} />
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
    const { quizOverview, quizPerformance, quizByDepartment, quizQuestionStats, quizTrend, quizTopPerformers } = this.state;

    const maxAttempts = Math.max(...quizTrend.map(t => t.attempts));

    const difficultyColor = (d: string): string => {
      switch (d) {
        case 'Easy': return '#10b981';
        case 'Medium': return '#f59e0b';
        case 'Hard': return '#ef4444';
        case 'Expert': return '#7c3aed';
        default: return '#64748b';
      }
    };

    return (
      <div className={styles.quizTab}>
        {/* Overview KPIs */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Quiz Overview</h3></div>
            <div className={styles.kpiGrid}>
              <div className={`${styles.kpiCard} ${styles.kpiPrimary}`}>
                <div className={styles.kpiValue}>{quizOverview.totalQuizzes}</div>
                <div className={styles.kpiLabel}>Total Quizzes</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{quizOverview.activeQuizzes}</div>
                <div className={styles.kpiLabel}>Active Quizzes</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{quizOverview.totalAttempts.toLocaleString()}</div>
                <div className={styles.kpiLabel}>Total Attempts</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{quizOverview.avgScore}%</div>
                <div className={styles.kpiLabel}>Avg Score</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{quizOverview.passRate}%</div>
                <div className={styles.kpiLabel}>Pass Rate</div>
              </div>
              <div className={styles.kpiCard}>
                <div className={styles.kpiValue}>{quizOverview.avgCompletionTime}</div>
                <div className={styles.kpiLabel}>Avg Completion Time</div>
              </div>
            </div>
          </div>
        </div>

        {/* Quiz Attempts Trend */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.sectionHeader}><h3>Quiz Attempts Trend <small>(12 Months)</small></h3></div>
            <div className={styles.trendChart}>
              <div className={styles.trendBars}>
                {quizTrend.map((t, i) => (
                  <div key={i} className={styles.trendBarCol}>
                    <div className={styles.trendBarValue}>{t.attempts}</div>
                    <div className={styles.trendBar} style={{ height: `${(t.attempts / maxAttempts) * 160}px` }} />
                    <div className={styles.trendBarLabel}>{t.month}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Quiz Performance Table */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Quiz Performance</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Quiz</th>
                  <th>Attempts</th>
                  <th>Avg Score</th>
                  <th>Pass Rate</th>
                  <th>Avg Time</th>
                  <th>Difficulty</th>
                </tr>
              </thead>
              <tbody>
                {quizPerformance.map((quiz, i) => (
                  <tr key={i}>
                    <td className={styles.cellTitle}>{quiz.title}</td>
                    <td className={styles.cellNumber}>{quiz.attempts}</td>
                    <td>
                      <span className={styles.rateBar}>
                        <span className={styles.rateBarFill} style={{ width: `${quiz.avgScore}%`, background: quiz.avgScore >= 80 ? '#10b981' : quiz.avgScore >= 60 ? '#f59e0b' : '#ef4444' }} />
                      </span>
                      {quiz.avgScore}%
                    </td>
                    <td className={quiz.passRate >= 85 ? styles.textSuccess : quiz.passRate >= 75 ? styles.textWarning : styles.textDanger}>
                      {quiz.passRate}%
                    </td>
                    <td>{quiz.avgTime}</td>
                    <td>
                      <span style={{ fontSize: 10, padding: '2px 8px', borderRadius: 4, backgroundColor: `${difficultyColor(quiz.difficulty)}15`, color: difficultyColor(quiz.difficulty) }}>
                        {quiz.difficulty}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Department Performance + Top Performers */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgAlt}`}>
            <div className={styles.twoColumnSection}>
              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Performance by Department</h3></div>
                <table className={styles.dataTable}>
                  <thead>
                    <tr>
                      <th>Department</th>
                      <th>Attempts</th>
                      <th>Avg Score</th>
                      <th>Pass Rate</th>
                      <th>Completion</th>
                    </tr>
                  </thead>
                  <tbody>
                    {quizByDepartment.map((dept, i) => (
                      <tr key={i}>
                        <td className={styles.cellTitle}>{dept.department}</td>
                        <td className={styles.cellNumber}>{dept.attempts}</td>
                        <td>{dept.avgScore}%</td>
                        <td className={dept.passRate >= 85 ? styles.textSuccess : dept.passRate >= 75 ? styles.textWarning : styles.textDanger}>
                          {dept.passRate}%
                        </td>
                        <td>
                          <span className={styles.rateBar}>
                            <span className={styles.rateBarFill} style={{ width: `${dept.completionRate}%`, background: '#0d9488' }} />
                          </span>
                          {dept.completionRate}%
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className={styles.columnPanel}>
                <div className={styles.sectionHeader}><h3>Top Performers</h3></div>
                <table className={styles.dataTable}>
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>Name</th>
                      <th>Dept</th>
                      <th>Quizzes</th>
                      <th>Avg Score</th>
                      <th>Perfect</th>
                    </tr>
                  </thead>
                  <tbody>
                    {quizTopPerformers.map((p, i) => (
                      <tr key={i}>
                        <td className={styles.cellRank}>{i + 1}</td>
                        <td className={styles.cellTitle}>{p.name}</td>
                        <td>{p.department}</td>
                        <td className={styles.cellNumber}>{p.quizzesCompleted}</td>
                        <td className={styles.textSuccess}>{p.avgScore}%</td>
                        <td className={styles.cellNumber}>{p.perfectScores}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>

        {/* Question Difficulty Analysis */}
        <div className={styles.section}>
          <div className={`${styles.sectionInner} ${styles.sectionBgWhite}`}>
            <div className={styles.sectionHeader}><h3>Question Difficulty Analysis</h3></div>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th>Question</th>
                  <th>Quiz</th>
                  <th>Correct Rate</th>
                  <th>Avg Time</th>
                  <th>Difficulty</th>
                </tr>
              </thead>
              <tbody>
                {quizQuestionStats.map((q, i) => (
                  <tr key={i}>
                    <td className={styles.cellTitle}>{q.question}</td>
                    <td>{q.quizTitle}</td>
                    <td>
                      <span className={styles.rateBar}>
                        <span className={styles.rateBarFill} style={{ width: `${q.correctRate}%`, background: q.correctRate >= 80 ? '#10b981' : q.correctRate >= 60 ? '#f59e0b' : '#ef4444' }} />
                      </span>
                      {q.correctRate}%
                    </td>
                    <td>{q.avgTime}</td>
                    <td>
                      <span style={{ fontSize: 10, padding: '2px 8px', borderRadius: 4, backgroundColor: `${difficultyColor(q.difficulty)}15`, color: difficultyColor(q.difficulty) }}>
                        {q.difficulty}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }
}
