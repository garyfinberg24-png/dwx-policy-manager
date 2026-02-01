/**
 * South African Demo Data for Protea Financial Services
 * Realistic sample data for customer demonstrations of the Policy Manager application.
 * All dates are relative to January 2026.
 */

// #region Interfaces

export interface IDemoEmployee {
  id: number;
  name: string;
  email: string;
  department: string;
  jobTitle: string;
  office: string;
  manager: string;
  startDate: Date;
}

export interface IDemoPolicy {
  id: number;
  title: string;
  policyNumber: string;
  category: string;
  status: 'Published' | 'Draft' | 'In Review' | 'Archived';
  version: string;
  effectiveDate: Date;
  reviewDate: Date;
  owner: string;
  department: string;
  requiresQuiz: boolean;
  requiresAcknowledgement: boolean;
  description: string;
}

export interface IDemoPolicyPack {
  id: number;
  name: string;
  description: string;
  type: string;
  policyIds: number[];
  targetGroups: string[];
}

export interface IDemoCampaign {
  id: number;
  campaignName: string;
  contentType: 'Policy' | 'Policy Pack';
  policyTitle: string;
  policyId: number;
  policyPackName?: string;
  policyPackId?: number;
  scope: string;
  targetUsers: string[];
  targetGroups: string[];
  status: 'Draft' | 'Scheduled' | 'Active' | 'Completed' | 'Paused';
  scheduledDate?: Date;
  distributedDate?: Date;
  dueDate?: Date;
  targetCount: number;
  totalSent: number;
  totalDelivered: number;
  totalOpened: number;
  totalAcknowledged: number;
  totalOverdue: number;
  totalExempted: number;
  totalFailed: number;
  escalationEnabled: boolean;
  reminderSchedule: string;
  isActive: boolean;
  completedDate?: Date;
  createdDate: Date;
  createdBy: string;
}

export interface IDemoAcknowledgement {
  id: number;
  employeeId: number;
  employeeName: string;
  policyId: number;
  policyTitle: string;
  department: string;
  status: 'Pending' | 'Sent' | 'Delivered' | 'Opened' | 'Acknowledged' | 'Overdue' | 'Exempted' | 'Failed';
  assignedDate: Date;
  dueDate: Date;
  sentDate?: Date;
  openedDate?: Date;
  acknowledgedDate?: Date;
  quizScore?: number;
  quizPassed?: boolean;
}

export interface IDemoDelegation {
  id: number;
  policyTitle: string;
  assignedTo: string;
  assignedBy: string;
  taskType: 'Draft' | 'Review' | 'Distribute' | 'Approve';
  priority: 'Low' | 'Medium' | 'High' | 'Critical';
  status: 'Pending' | 'InProgress' | 'Completed' | 'Overdue';
  assignedDate: Date;
  dueDate: Date;
  completedDate?: Date;
  notes: string;
}

export interface IDemoTeamMember {
  id: number;
  name: string;
  email: string;
  department: string;
  policiesAssigned: number;
  policiesAcknowledged: number;
  policiesPending: number;
  policiesOverdue: number;
  compliancePercent: number;
  lastActivity: Date;
}

export interface IDemoSlaMetric {
  type: string;
  targetDays: number;
  actualAvgDays: number;
  percentMet: number;
  status: 'Met' | 'At Risk' | 'Breached';
  breaches: Array<{
    policyTitle: string;
    targetDays: number;
    actualDays: number;
    breachDate: Date;
    department: string;
  }>;
}

export interface IDemoAuditEntry {
  id: number;
  timestamp: Date;
  user: string;
  action: string;
  policyTitle: string;
  details: string;
  ipAddress: string;
}

export interface IDemoViolation {
  id: number;
  severity: 'Low' | 'Medium' | 'High' | 'Critical';
  policyTitle: string;
  department: string;
  description: string;
  status: 'Open' | 'Under Investigation' | 'Resolved' | 'Escalated';
  reportedDate: Date;
  resolvedDate?: Date;
}

export interface IDemoQuizResult {
  id: number;
  employeeName: string;
  department: string;
  policyTitle: string;
  score: number;
  passed: boolean;
  attemptDate: Date;
  attemptNumber: number;
  timeTaken: number;
}

export interface IDemoDataSummary {
  totalEmployees: number;
  totalPolicies: number;
  publishedPolicies: number;
  draftPolicies: number;
  overallComplianceRate: number;
  overdueCount: number;
  activeCampaigns: number;
  completedCampaigns: number;
  totalAcknowledgements: number;
  acknowledgedCount: number;
  pendingCount: number;
  openViolations: number;
  averageQuizScore: number;
  departmentsTracked: number;
}

// #endregion Interfaces

// #region Employees

/** 42 employees across all departments at Protea Financial Services */
export const DEMO_EMPLOYEES: IDemoEmployee[] = [
  { id: 1, name: 'Sipho Dlamini', email: 'sipho.dlamini@proteafs.co.za', department: 'IT', jobTitle: 'Chief Information Officer', office: 'Johannesburg', manager: 'Pieter van der Merwe', startDate: new Date(2018, 2, 1) },
  { id: 2, name: 'Pieter van der Merwe', email: 'pieter.vandermerwe@proteafs.co.za', department: 'Executive', jobTitle: 'Managing Director', office: 'Johannesburg', manager: '', startDate: new Date(2015, 0, 12) },
  { id: 3, name: 'Naledi Mokoena', email: 'naledi.mokoena@proteafs.co.za', department: 'HR', jobTitle: 'HR Director', office: 'Johannesburg', manager: 'Pieter van der Merwe', startDate: new Date(2017, 5, 15) },
  { id: 4, name: 'Fatima Patel', email: 'fatima.patel@proteafs.co.za', department: 'Finance', jobTitle: 'Financial Director', office: 'Johannesburg', manager: 'Pieter van der Merwe', startDate: new Date(2016, 8, 1) },
  { id: 5, name: 'Johan Botha', email: 'johan.botha@proteafs.co.za', department: 'Legal', jobTitle: 'Head of Legal', office: 'Cape Town', manager: 'Pieter van der Merwe', startDate: new Date(2017, 1, 20) },
  { id: 6, name: 'Thandiwe Nkosi', email: 'thandiwe.nkosi@proteafs.co.za', department: 'Compliance', jobTitle: 'Chief Compliance Officer', office: 'Johannesburg', manager: 'Pieter van der Merwe', startDate: new Date(2019, 3, 1) },
  { id: 7, name: 'Rajesh Naidoo', email: 'rajesh.naidoo@proteafs.co.za', department: 'IT', jobTitle: 'IT Manager', office: 'Durban', manager: 'Sipho Dlamini', startDate: new Date(2019, 7, 12) },
  { id: 8, name: 'Lindiwe Zulu', email: 'lindiwe.zulu@proteafs.co.za', department: 'HR', jobTitle: 'HR Business Partner', office: 'Johannesburg', manager: 'Naledi Mokoena', startDate: new Date(2020, 0, 6) },
  { id: 9, name: 'Charl Pretorius', email: 'charl.pretorius@proteafs.co.za', department: 'Finance', jobTitle: 'Senior Financial Analyst', office: 'Johannesburg', manager: 'Fatima Patel', startDate: new Date(2020, 4, 18) },
  { id: 10, name: 'Nomvula Khumalo', email: 'nomvula.khumalo@proteafs.co.za', department: 'Operations', jobTitle: 'Operations Director', office: 'Johannesburg', manager: 'Pieter van der Merwe', startDate: new Date(2018, 10, 1) },
  { id: 11, name: 'Craig Williams', email: 'craig.williams@proteafs.co.za', department: 'Sales', jobTitle: 'Head of Sales', office: 'Cape Town', manager: 'Pieter van der Merwe', startDate: new Date(2019, 1, 15) },
  { id: 12, name: 'Zanele Mthembu', email: 'zanele.mthembu@proteafs.co.za', department: 'Marketing', jobTitle: 'Marketing Manager', office: 'Johannesburg', manager: 'Craig Williams', startDate: new Date(2021, 2, 1) },
  { id: 13, name: 'Hendrik Viljoen', email: 'hendrik.viljoen@proteafs.co.za', department: 'Engineering', jobTitle: 'Head of Engineering', office: 'Cape Town', manager: 'Sipho Dlamini', startDate: new Date(2019, 6, 8) },
  { id: 14, name: 'Priya Govender', email: 'priya.govender@proteafs.co.za', department: 'Compliance', jobTitle: 'Compliance Analyst', office: 'Durban', manager: 'Thandiwe Nkosi', startDate: new Date(2021, 8, 1) },
  { id: 15, name: 'Thabo Molefe', email: 'thabo.molefe@proteafs.co.za', department: 'IT', jobTitle: 'Systems Administrator', office: 'Johannesburg', manager: 'Rajesh Naidoo', startDate: new Date(2022, 0, 10) },
  { id: 16, name: 'Annemarie du Plessis', email: 'annemarie.duplessis@proteafs.co.za', department: 'Finance', jobTitle: 'Accountant', office: 'Pretoria', manager: 'Fatima Patel', startDate: new Date(2021, 5, 14) },
  { id: 17, name: 'Bongani Sithole', email: 'bongani.sithole@proteafs.co.za', department: 'Sales', jobTitle: 'Sales Executive', office: 'Durban', manager: 'Craig Williams', startDate: new Date(2022, 3, 1) },
  { id: 18, name: 'Michelle September', email: 'michelle.september@proteafs.co.za', department: 'HR', jobTitle: 'Recruitment Specialist', office: 'Cape Town', manager: 'Naledi Mokoena', startDate: new Date(2022, 7, 22) },
  { id: 19, name: 'Vuyo Madonsela', email: 'vuyo.madonsela@proteafs.co.za', department: 'Legal', jobTitle: 'Legal Advisor', office: 'Johannesburg', manager: 'Johan Botha', startDate: new Date(2021, 11, 1) },
  { id: 20, name: 'Riyaad Jacobs', email: 'riyaad.jacobs@proteafs.co.za', department: 'Facilities', jobTitle: 'Facilities Manager', office: 'Cape Town', manager: 'Nomvula Khumalo', startDate: new Date(2020, 9, 5) },
  { id: 21, name: 'Lerato Mahlangu', email: 'lerato.mahlangu@proteafs.co.za', department: 'Procurement', jobTitle: 'Procurement Manager', office: 'Johannesburg', manager: 'Nomvula Khumalo', startDate: new Date(2021, 1, 15) },
  { id: 22, name: 'Suresh Pillay', email: 'suresh.pillay@proteafs.co.za', department: 'IT', jobTitle: 'Software Developer', office: 'Durban', manager: 'Hendrik Viljoen', startDate: new Date(2023, 0, 9) },
  { id: 23, name: 'Ayanda Zwane', email: 'ayanda.zwane@proteafs.co.za', department: 'Operations', jobTitle: 'Operations Analyst', office: 'Johannesburg', manager: 'Nomvula Khumalo', startDate: new Date(2023, 4, 1) },
  { id: 24, name: 'Karen Mostert', email: 'karen.mostert@proteafs.co.za', department: 'Finance', jobTitle: 'Financial Controller', office: 'Johannesburg', manager: 'Fatima Patel', startDate: new Date(2020, 2, 16) },
  { id: 25, name: 'Tshepo Mabaso', email: 'tshepo.mabaso@proteafs.co.za', department: 'Sales', jobTitle: 'Sales Representative', office: 'Pretoria', manager: 'Craig Williams', startDate: new Date(2023, 6, 1) },
  { id: 26, name: 'Ncumisa Dyani', email: 'ncumisa.dyani@proteafs.co.za', department: 'Compliance', jobTitle: 'Risk Analyst', office: 'Cape Town', manager: 'Thandiwe Nkosi', startDate: new Date(2022, 10, 7) },
  { id: 27, name: 'Willem Erasmus', email: 'willem.erasmus@proteafs.co.za', department: 'Engineering', jobTitle: 'Senior Developer', office: 'Cape Town', manager: 'Hendrik Viljoen', startDate: new Date(2021, 3, 12) },
  { id: 28, name: 'Nozipho Buthelezi', email: 'nozipho.buthelezi@proteafs.co.za', department: 'Marketing', jobTitle: 'Digital Marketing Specialist', office: 'Johannesburg', manager: 'Zanele Mthembu', startDate: new Date(2023, 8, 1) },
  { id: 29, name: 'Imraan Moosa', email: 'imraan.moosa@proteafs.co.za', department: 'Legal', jobTitle: 'Contract Specialist', office: 'Durban', manager: 'Johan Botha', startDate: new Date(2022, 5, 20) },
  { id: 30, name: 'Marike Joubert', email: 'marike.joubert@proteafs.co.za', department: 'HR', jobTitle: 'Learning & Development Specialist', office: 'Pretoria', manager: 'Naledi Mokoena', startDate: new Date(2023, 1, 1) },
  { id: 31, name: 'Sibusiso Ndlovu', email: 'sibusiso.ndlovu@proteafs.co.za', department: 'IT', jobTitle: 'Network Engineer', office: 'Johannesburg', manager: 'Rajesh Naidoo', startDate: new Date(2022, 8, 12) },
  { id: 32, name: 'Yolande van Wyk', email: 'yolande.vanwyk@proteafs.co.za', department: 'Finance', jobTitle: 'Credit Analyst', office: 'Cape Town', manager: 'Fatima Patel', startDate: new Date(2023, 3, 15) },
  { id: 33, name: 'Mandla Cele', email: 'mandla.cele@proteafs.co.za', department: 'Operations', jobTitle: 'Logistics Coordinator', office: 'Durban', manager: 'Nomvula Khumalo', startDate: new Date(2024, 0, 8) },
  { id: 34, name: 'Samantha Adams', email: 'samantha.adams@proteafs.co.za', department: 'Sales', jobTitle: 'Business Development Manager', office: 'Cape Town', manager: 'Craig Williams', startDate: new Date(2023, 10, 1) },
  { id: 35, name: 'Kamohelo Tlali', email: 'kamohelo.tlali@proteafs.co.za', department: 'Engineering', jobTitle: 'DevOps Engineer', office: 'Johannesburg', manager: 'Hendrik Viljoen', startDate: new Date(2024, 2, 18) },
  { id: 36, name: 'Faizel Davids', email: 'faizel.davids@proteafs.co.za', department: 'Procurement', jobTitle: 'Procurement Officer', office: 'Cape Town', manager: 'Lerato Mahlangu', startDate: new Date(2024, 5, 1) },
  { id: 37, name: 'Dineo Masemola', email: 'dineo.masemola@proteafs.co.za', department: 'Compliance', jobTitle: 'POPIA Officer', office: 'Johannesburg', manager: 'Thandiwe Nkosi', startDate: new Date(2024, 1, 12) },
  { id: 38, name: 'Grant Thompson', email: 'grant.thompson@proteafs.co.za', department: 'Facilities', jobTitle: 'Health & Safety Officer', office: 'Johannesburg', manager: 'Riyaad Jacobs', startDate: new Date(2024, 7, 5) },
  { id: 39, name: 'Zinhle Ngcobo', email: 'zinhle.ngcobo@proteafs.co.za', department: 'Marketing', jobTitle: 'Content Creator', office: 'Durban', manager: 'Zanele Mthembu', startDate: new Date(2024, 9, 1) },
  { id: 40, name: 'Anika Liebenberg', email: 'anika.liebenberg@proteafs.co.za', department: 'Legal', jobTitle: 'Paralegal', office: 'Pretoria', manager: 'Johan Botha', startDate: new Date(2025, 0, 13) },
  { id: 41, name: 'Lungelo Mkhize', email: 'lungelo.mkhize@proteafs.co.za', department: 'Sales', jobTitle: 'Account Manager', office: 'Johannesburg', manager: 'Craig Williams', startDate: new Date(2025, 3, 1) },
  { id: 42, name: 'Charlize Steyn', email: 'charlize.steyn@proteafs.co.za', department: 'IT', jobTitle: 'Business Analyst', office: 'Pretoria', manager: 'Sipho Dlamini', startDate: new Date(2025, 5, 16) },
];

// #endregion Employees

// #region Policies

/** 18 policies covering SA-specific regulatory and corporate governance requirements */
export const DEMO_POLICIES: IDemoPolicy[] = [
  { id: 1, title: 'POPIA Data Privacy Policy', policyNumber: 'PFS-POPIA-001', category: 'Data Privacy', status: 'Published', version: '3.1', effectiveDate: new Date(2025, 6, 1), reviewDate: new Date(2026, 6, 1), owner: 'Thandiwe Nkosi', department: 'Compliance', requiresQuiz: true, requiresAcknowledgement: true, description: 'Protection of Personal Information Act compliance policy governing the collection, processing, storage, and sharing of personal information.' },
  { id: 2, title: 'BBBEE Compliance Policy', policyNumber: 'PFS-BBBEE-001', category: 'Regulatory', status: 'Published', version: '2.4', effectiveDate: new Date(2025, 3, 1), reviewDate: new Date(2026, 3, 1), owner: 'Naledi Mokoena', department: 'HR', requiresQuiz: false, requiresAcknowledgement: true, description: 'Broad-Based Black Economic Empowerment policy ensuring transformation targets and preferential procurement compliance.' },
  { id: 3, title: 'Employment Equity Policy', policyNumber: 'PFS-EEA-001', category: 'Employment', status: 'Published', version: '2.2', effectiveDate: new Date(2025, 0, 15), reviewDate: new Date(2026, 0, 15), owner: 'Naledi Mokoena', department: 'HR', requiresQuiz: true, requiresAcknowledgement: true, description: 'Employment Equity Act compliance policy promoting equal opportunity and fair treatment through elimination of unfair discrimination.' },
  { id: 4, title: 'Occupational Health & Safety Policy', policyNumber: 'PFS-OHSA-001', category: 'Health & Safety', status: 'Published', version: '4.0', effectiveDate: new Date(2025, 8, 1), reviewDate: new Date(2026, 8, 1), owner: 'Grant Thompson', department: 'Facilities', requiresQuiz: true, requiresAcknowledgement: true, description: 'Occupational Health and Safety Act policy ensuring a safe and healthy working environment for all employees and contractors.' },
  { id: 5, title: 'King IV Corporate Governance Policy', policyNumber: 'PFS-KIV-001', category: 'Governance', status: 'Published', version: '1.8', effectiveDate: new Date(2025, 4, 1), reviewDate: new Date(2026, 4, 1), owner: 'Johan Botha', department: 'Legal', requiresQuiz: false, requiresAcknowledgement: true, description: 'Corporate governance policy aligned with King IV principles covering ethical leadership, stakeholder inclusivity, and integrated reporting.' },
  { id: 6, title: 'FICA & Anti-Money Laundering Policy', policyNumber: 'PFS-FICA-001', category: 'Financial Compliance', status: 'Published', version: '3.5', effectiveDate: new Date(2025, 1, 1), reviewDate: new Date(2026, 1, 1), owner: 'Fatima Patel', department: 'Finance', requiresQuiz: true, requiresAcknowledgement: true, description: 'Financial Intelligence Centre Act compliance policy covering client identification, record-keeping, and suspicious transaction reporting.' },
  { id: 7, title: 'Information Security Policy', policyNumber: 'PFS-ISP-001', category: 'IT Security', status: 'Published', version: '5.2', effectiveDate: new Date(2025, 5, 1), reviewDate: new Date(2026, 5, 1), owner: 'Sipho Dlamini', department: 'IT', requiresQuiz: true, requiresAcknowledgement: true, description: 'Comprehensive information security policy covering data classification, access control, incident response, and cybersecurity measures.' },
  { id: 8, title: 'Companies Act Compliance Policy', policyNumber: 'PFS-CA-001', category: 'Regulatory', status: 'Published', version: '2.0', effectiveDate: new Date(2025, 2, 1), reviewDate: new Date(2026, 2, 1), owner: 'Johan Botha', department: 'Legal', requiresQuiz: false, requiresAcknowledgement: true, description: 'Companies Act 71 of 2008 compliance framework covering director duties, financial reporting, and shareholder rights.' },
  { id: 9, title: 'Anti-Bribery & Corruption Policy', policyNumber: 'PFS-ABC-001', category: 'Ethics', status: 'Published', version: '2.1', effectiveDate: new Date(2025, 7, 1), reviewDate: new Date(2026, 7, 1), owner: 'Thandiwe Nkosi', department: 'Compliance', requiresQuiz: true, requiresAcknowledgement: true, description: 'Zero-tolerance policy against bribery, corruption, and facilitation payments in accordance with the Prevention and Combating of Corrupt Activities Act.' },
  { id: 10, title: 'Whistleblower Protection Policy', policyNumber: 'PFS-WBP-001', category: 'Ethics', status: 'Published', version: '1.5', effectiveDate: new Date(2025, 9, 1), reviewDate: new Date(2026, 9, 1), owner: 'Johan Botha', department: 'Legal', requiresQuiz: false, requiresAcknowledgement: true, description: 'Protected Disclosures Act aligned policy encouraging reporting of irregular conduct without fear of retaliation.' },
  { id: 11, title: 'Remote Work & Flexible Arrangements Policy', policyNumber: 'PFS-RWK-001', category: 'Employment', status: 'Published', version: '2.0', effectiveDate: new Date(2025, 10, 1), reviewDate: new Date(2026, 10, 1), owner: 'Naledi Mokoena', department: 'HR', requiresQuiz: false, requiresAcknowledgement: true, description: 'Guidelines for remote working, hybrid arrangements, and flexible working hours in compliance with the Basic Conditions of Employment Act.' },
  { id: 12, title: 'Disaster Recovery & Business Continuity Policy', policyNumber: 'PFS-DRP-001', category: 'IT Security', status: 'In Review', version: '3.0', effectiveDate: new Date(2025, 11, 1), reviewDate: new Date(2026, 5, 1), owner: 'Sipho Dlamini', department: 'IT', requiresQuiz: true, requiresAcknowledgement: true, description: 'Business continuity and disaster recovery procedures including load shedding contingency plans and data backup protocols.' },
  { id: 13, title: 'Procurement & Supply Chain Policy', policyNumber: 'PFS-PRC-001', category: 'Operations', status: 'Published', version: '1.6', effectiveDate: new Date(2025, 3, 15), reviewDate: new Date(2026, 3, 15), owner: 'Lerato Mahlangu', department: 'Procurement', requiresQuiz: false, requiresAcknowledgement: true, description: 'Procurement procedures ensuring BBBEE supplier development, preferential procurement targets, and ethical sourcing practices.' },
  { id: 14, title: 'Social Media & Communications Policy', policyNumber: 'PFS-SMC-001', category: 'Communications', status: 'Draft', version: '0.9', effectiveDate: new Date(2026, 1, 1), reviewDate: new Date(2026, 7, 1), owner: 'Zanele Mthembu', department: 'Marketing', requiresQuiz: false, requiresAcknowledgement: true, description: 'Guidelines governing employee use of social media, external communications, and brand representation.' },
  { id: 15, title: 'Environmental, Social & Governance Policy', policyNumber: 'PFS-ESG-001', category: 'Governance', status: 'In Review', version: '1.2', effectiveDate: new Date(2026, 0, 15), reviewDate: new Date(2026, 6, 15), owner: 'Thandiwe Nkosi', department: 'Compliance', requiresQuiz: false, requiresAcknowledgement: true, description: 'ESG commitment framework covering carbon footprint reduction, social impact, and sustainable business practices.' },
  { id: 16, title: 'Consumer Protection Act Compliance Policy', policyNumber: 'PFS-CPA-001', category: 'Regulatory', status: 'Published', version: '2.3', effectiveDate: new Date(2025, 5, 15), reviewDate: new Date(2026, 5, 15), owner: 'Johan Botha', department: 'Legal', requiresQuiz: true, requiresAcknowledgement: true, description: 'Policy ensuring compliance with the Consumer Protection Act covering fair marketing, product liability, and customer rights.' },
  { id: 17, title: 'Expense & Travel Policy', policyNumber: 'PFS-ETP-001', category: 'Finance', status: 'Published', version: '3.2', effectiveDate: new Date(2025, 0, 1), reviewDate: new Date(2026, 0, 1), owner: 'Fatima Patel', department: 'Finance', requiresQuiz: false, requiresAcknowledgement: true, description: 'Guidelines for business travel, expense claims, per diem rates, and reimbursement procedures.' },
  { id: 18, title: 'Code of Ethics & Business Conduct', policyNumber: 'PFS-COE-001', category: 'Ethics', status: 'Archived', version: '2.0', effectiveDate: new Date(2023, 0, 1), reviewDate: new Date(2025, 0, 1), owner: 'Pieter van der Merwe', department: 'Executive', requiresQuiz: true, requiresAcknowledgement: true, description: 'Superseded code of ethics replaced by updated Anti-Bribery & Corruption and Whistleblower policies.' },
];

// #endregion Policies

// #region Policy Packs

/** 6 policy packs grouping related policies for targeted distribution */
export const DEMO_POLICY_PACKS: IDemoPolicyPack[] = [
  { id: 1, name: 'New Employee Onboarding Pack', description: 'Essential policies every new employee must acknowledge within their first 30 days at Protea Financial Services.', type: 'Onboarding', policyIds: [1, 3, 4, 7, 9, 11], targetGroups: ['All Employees'] },
  { id: 2, name: 'Annual Regulatory Compliance Pack', description: 'Annual re-acknowledgement pack for all SA regulatory compliance policies including POPIA, FICA, and BBBEE.', type: 'Annual Review', policyIds: [1, 2, 6, 8, 16], targetGroups: ['All Employees'] },
  { id: 3, name: 'Finance & Risk Pack', description: 'Specialised compliance pack for Finance department covering financial regulations and anti-money laundering.', type: 'Department', policyIds: [6, 8, 9, 17], targetGroups: ['Finance', 'Compliance'] },
  { id: 4, name: 'IT Security Essentials Pack', description: 'Information security and disaster recovery policies for all technology staff.', type: 'Department', policyIds: [1, 7, 12], targetGroups: ['IT', 'Engineering'] },
  { id: 5, name: 'Management Governance Pack', description: 'Corporate governance and leadership policies for managers and directors aligned with King IV principles.', type: 'Role-Based', policyIds: [5, 8, 9, 10, 15], targetGroups: ['Executive', 'Management'] },
  { id: 6, name: 'Health & Safety Refresher Pack', description: 'Annual OHS refresher pack for all office locations covering workplace safety and emergency procedures.', type: 'Annual Review', policyIds: [4, 11], targetGroups: ['All Employees'] },
];

// #endregion Policy Packs

// #region Campaigns

/** 10 distribution campaigns in various states of completion */
export const DEMO_CAMPAIGNS: IDemoCampaign[] = [
  {
    id: 1, campaignName: 'Q1 2026 POPIA Awareness Campaign', contentType: 'Policy', policyTitle: 'POPIA Data Privacy Policy', policyId: 1,
    scope: 'Organisation-wide', targetUsers: [], targetGroups: ['All Employees'],
    status: 'Active', scheduledDate: new Date(2026, 0, 6), distributedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6),
    targetCount: 42, totalSent: 42, totalDelivered: 40, totalOpened: 35, totalAcknowledged: 28, totalOverdue: 7, totalExempted: 2, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: true, createdDate: new Date(2025, 11, 20), createdBy: 'Thandiwe Nkosi'
  },
  {
    id: 2, campaignName: 'Annual Regulatory Compliance 2026', contentType: 'Policy Pack', policyTitle: 'Annual Regulatory Compliance Pack', policyId: 1, policyPackName: 'Annual Regulatory Compliance Pack', policyPackId: 2,
    scope: 'Organisation-wide', targetUsers: [], targetGroups: ['All Employees'],
    status: 'Active', scheduledDate: new Date(2026, 0, 13), distributedDate: new Date(2026, 0, 13), dueDate: new Date(2026, 1, 28),
    targetCount: 42, totalSent: 42, totalDelivered: 41, totalOpened: 30, totalAcknowledged: 18, totalOverdue: 3, totalExempted: 1, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Bi-weekly', isActive: true, createdDate: new Date(2025, 11, 28), createdBy: 'Thandiwe Nkosi'
  },
  {
    id: 3, campaignName: 'Finance Team FICA Refresher', contentType: 'Policy Pack', policyTitle: 'Finance & Risk Pack', policyId: 6, policyPackName: 'Finance & Risk Pack', policyPackId: 3,
    scope: 'Department', targetUsers: [], targetGroups: ['Finance', 'Compliance'],
    status: 'Completed', scheduledDate: new Date(2025, 10, 1), distributedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1),
    targetCount: 10, totalSent: 10, totalDelivered: 10, totalOpened: 10, totalAcknowledged: 10, totalOverdue: 0, totalExempted: 0, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: false, completedDate: new Date(2025, 10, 25), createdDate: new Date(2025, 9, 20), createdBy: 'Fatima Patel'
  },
  {
    id: 4, campaignName: 'IT Security Policy Rollout', contentType: 'Policy Pack', policyTitle: 'IT Security Essentials Pack', policyId: 7, policyPackName: 'IT Security Essentials Pack', policyPackId: 4,
    scope: 'Department', targetUsers: [], targetGroups: ['IT', 'Engineering'],
    status: 'Active', scheduledDate: new Date(2026, 0, 8), distributedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31),
    targetCount: 9, totalSent: 9, totalDelivered: 9, totalOpened: 8, totalAcknowledged: 6, totalOverdue: 1, totalExempted: 0, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: true, createdDate: new Date(2025, 11, 30), createdBy: 'Sipho Dlamini'
  },
  {
    id: 5, campaignName: 'New Joiner Onboarding - Jan 2026', contentType: 'Policy Pack', policyTitle: 'New Employee Onboarding Pack', policyId: 1, policyPackName: 'New Employee Onboarding Pack', policyPackId: 1,
    scope: 'Targeted', targetUsers: ['lungelo.mkhize@proteafs.co.za', 'charlize.steyn@proteafs.co.za'], targetGroups: [],
    status: 'Completed', scheduledDate: new Date(2025, 11, 15), distributedDate: new Date(2025, 11, 15), dueDate: new Date(2026, 0, 15),
    targetCount: 2, totalSent: 2, totalDelivered: 2, totalOpened: 2, totalAcknowledged: 2, totalOverdue: 0, totalExempted: 0, totalFailed: 0,
    escalationEnabled: false, reminderSchedule: 'Daily', isActive: false, completedDate: new Date(2026, 0, 10), createdDate: new Date(2025, 11, 12), createdBy: 'Naledi Mokoena'
  },
  {
    id: 6, campaignName: 'Anti-Bribery Annual Refresher', contentType: 'Policy', policyTitle: 'Anti-Bribery & Corruption Policy', policyId: 9,
    scope: 'Organisation-wide', targetUsers: [], targetGroups: ['All Employees'],
    status: 'Completed', scheduledDate: new Date(2025, 9, 1), distributedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1),
    targetCount: 42, totalSent: 42, totalDelivered: 41, totalOpened: 40, totalAcknowledged: 39, totalOverdue: 0, totalExempted: 2, totalFailed: 1,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: false, completedDate: new Date(2025, 9, 28), createdDate: new Date(2025, 8, 15), createdBy: 'Thandiwe Nkosi'
  },
  {
    id: 7, campaignName: 'OHS Policy Update Distribution', contentType: 'Policy', policyTitle: 'Occupational Health & Safety Policy', policyId: 4,
    scope: 'Organisation-wide', targetUsers: [], targetGroups: ['All Employees'],
    status: 'Scheduled', scheduledDate: new Date(2026, 1, 3), dueDate: new Date(2026, 2, 3),
    targetCount: 42, totalSent: 0, totalDelivered: 0, totalOpened: 0, totalAcknowledged: 0, totalOverdue: 0, totalExempted: 0, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: false, createdDate: new Date(2026, 0, 20), createdBy: 'Grant Thompson'
  },
  {
    id: 8, campaignName: 'Management Governance Pack 2026', contentType: 'Policy Pack', policyTitle: 'Management Governance Pack', policyId: 5, policyPackName: 'Management Governance Pack', policyPackId: 5,
    scope: 'Role-Based', targetUsers: [], targetGroups: ['Executive', 'Management'],
    status: 'Active', scheduledDate: new Date(2026, 0, 15), distributedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 1, 15),
    targetCount: 12, totalSent: 12, totalDelivered: 12, totalOpened: 10, totalAcknowledged: 7, totalOverdue: 2, totalExempted: 0, totalFailed: 0,
    escalationEnabled: true, reminderSchedule: 'Weekly', isActive: true, createdDate: new Date(2026, 0, 10), createdBy: 'Johan Botha'
  },
  {
    id: 9, campaignName: 'ESG Policy Draft Review', contentType: 'Policy', policyTitle: 'Environmental, Social & Governance Policy', policyId: 15,
    scope: 'Targeted', targetUsers: ['thandiwe.nkosi@proteafs.co.za', 'johan.botha@proteafs.co.za', 'naledi.mokoena@proteafs.co.za', 'fatima.patel@proteafs.co.za'], targetGroups: [],
    status: 'Paused', scheduledDate: new Date(2026, 0, 20), distributedDate: new Date(2026, 0, 20), dueDate: new Date(2026, 1, 20),
    targetCount: 4, totalSent: 4, totalDelivered: 4, totalOpened: 3, totalAcknowledged: 1, totalOverdue: 0, totalExempted: 0, totalFailed: 0,
    escalationEnabled: false, reminderSchedule: 'None', isActive: false, createdDate: new Date(2026, 0, 18), createdBy: 'Thandiwe Nkosi'
  },
  {
    id: 10, campaignName: 'Employment Equity Awareness - Sales', contentType: 'Policy', policyTitle: 'Employment Equity Policy', policyId: 3,
    scope: 'Department', targetUsers: [], targetGroups: ['Sales'],
    status: 'Draft', dueDate: new Date(2026, 2, 1),
    targetCount: 5, totalSent: 0, totalDelivered: 0, totalOpened: 0, totalAcknowledged: 0, totalOverdue: 0, totalExempted: 0, totalFailed: 0,
    escalationEnabled: false, reminderSchedule: 'Weekly', isActive: false, createdDate: new Date(2026, 0, 28), createdBy: 'Naledi Mokoena'
  },
];

// #endregion Campaigns

// #region Acknowledgements

/** 65 acknowledgement records tracking employee policy acknowledgement status */
export const DEMO_ACKNOWLEDGEMENTS: IDemoAcknowledgement[] = [
  // POPIA Policy - Campaign 1 (various statuses to show realistic distribution)
  { id: 1, employeeId: 1, employeeName: 'Sipho Dlamini', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 7), acknowledgedDate: new Date(2026, 0, 7), quizScore: 95, quizPassed: true },
  { id: 2, employeeId: 3, employeeName: 'Naledi Mokoena', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 6), acknowledgedDate: new Date(2026, 0, 8), quizScore: 90, quizPassed: true },
  { id: 3, employeeId: 4, employeeName: 'Fatima Patel', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 7), acknowledgedDate: new Date(2026, 0, 9), quizScore: 100, quizPassed: true },
  { id: 4, employeeId: 6, employeeName: 'Thandiwe Nkosi', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Compliance', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 6), acknowledgedDate: new Date(2026, 0, 6), quizScore: 100, quizPassed: true },
  { id: 5, employeeId: 7, employeeName: 'Rajesh Naidoo', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 8), acknowledgedDate: new Date(2026, 0, 10), quizScore: 85, quizPassed: true },
  { id: 6, employeeId: 9, employeeName: 'Charl Pretorius', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 7), acknowledgedDate: new Date(2026, 0, 8), quizScore: 90, quizPassed: true },
  { id: 7, employeeId: 11, employeeName: 'Craig Williams', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', status: 'Opened', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 15) },
  { id: 8, employeeId: 12, employeeName: 'Zanele Mthembu', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Marketing', status: 'Delivered', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6) },
  { id: 9, employeeId: 17, employeeName: 'Bongani Sithole', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', status: 'Overdue', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 0, 20), sentDate: new Date(2026, 0, 6) },
  { id: 10, employeeId: 25, employeeName: 'Tshepo Mabaso', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', status: 'Overdue', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 0, 20), sentDate: new Date(2026, 0, 6), openedDate: new Date(2026, 0, 12) },
  { id: 11, employeeId: 34, employeeName: 'Samantha Adams', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', status: 'Sent', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6) },
  { id: 12, employeeId: 41, employeeName: 'Lungelo Mkhize', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', status: 'Exempted', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6) },

  // FICA Policy - Finance team (high compliance)
  { id: 13, employeeId: 4, employeeName: 'Fatima Patel', policyId: 6, policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1), sentDate: new Date(2025, 10, 1), openedDate: new Date(2025, 10, 1), acknowledgedDate: new Date(2025, 10, 2), quizScore: 100, quizPassed: true },
  { id: 14, employeeId: 9, employeeName: 'Charl Pretorius', policyId: 6, policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1), sentDate: new Date(2025, 10, 1), openedDate: new Date(2025, 10, 2), acknowledgedDate: new Date(2025, 10, 3), quizScore: 95, quizPassed: true },
  { id: 15, employeeId: 16, employeeName: 'Annemarie du Plessis', policyId: 6, policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1), sentDate: new Date(2025, 10, 1), openedDate: new Date(2025, 10, 1), acknowledgedDate: new Date(2025, 10, 4), quizScore: 88, quizPassed: true },
  { id: 16, employeeId: 24, employeeName: 'Karen Mostert', policyId: 6, policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1), sentDate: new Date(2025, 10, 1), openedDate: new Date(2025, 10, 2), acknowledgedDate: new Date(2025, 10, 5), quizScore: 92, quizPassed: true },
  { id: 17, employeeId: 32, employeeName: 'Yolande van Wyk', policyId: 6, policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 11, 1), sentDate: new Date(2025, 10, 1), openedDate: new Date(2025, 10, 3), acknowledgedDate: new Date(2025, 10, 5), quizScore: 85, quizPassed: true },

  // Information Security - IT team
  { id: 18, employeeId: 1, employeeName: 'Sipho Dlamini', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 8), acknowledgedDate: new Date(2026, 0, 9), quizScore: 100, quizPassed: true },
  { id: 19, employeeId: 7, employeeName: 'Rajesh Naidoo', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 9), acknowledgedDate: new Date(2026, 0, 10), quizScore: 92, quizPassed: true },
  { id: 20, employeeId: 15, employeeName: 'Thabo Molefe', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 10), acknowledgedDate: new Date(2026, 0, 12), quizScore: 88, quizPassed: true },
  { id: 21, employeeId: 22, employeeName: 'Suresh Pillay', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Opened', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 15) },
  { id: 22, employeeId: 31, employeeName: 'Sibusiso Ndlovu', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 9), acknowledgedDate: new Date(2026, 0, 11), quizScore: 95, quizPassed: true },
  { id: 23, employeeId: 42, employeeName: 'Charlize Steyn', policyId: 7, policyTitle: 'Information Security Policy', department: 'IT', status: 'Delivered', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8) },
  { id: 24, employeeId: 13, employeeName: 'Hendrik Viljoen', policyId: 7, policyTitle: 'Information Security Policy', department: 'Engineering', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 8), acknowledgedDate: new Date(2026, 0, 9), quizScore: 90, quizPassed: true },
  { id: 25, employeeId: 27, employeeName: 'Willem Erasmus', policyId: 7, policyTitle: 'Information Security Policy', department: 'Engineering', status: 'Overdue', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 22), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 14) },
  { id: 26, employeeId: 35, employeeName: 'Kamohelo Tlali', policyId: 7, policyTitle: 'Information Security Policy', department: 'Engineering', status: 'Acknowledged', assignedDate: new Date(2026, 0, 8), dueDate: new Date(2026, 0, 31), sentDate: new Date(2026, 0, 8), openedDate: new Date(2026, 0, 10), acknowledgedDate: new Date(2026, 0, 13), quizScore: 80, quizPassed: true },

  // Employment Equity - broad distribution
  { id: 27, employeeId: 3, employeeName: 'Naledi Mokoena', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1), openedDate: new Date(2025, 11, 1), acknowledgedDate: new Date(2025, 11, 2), quizScore: 100, quizPassed: true },
  { id: 28, employeeId: 8, employeeName: 'Lindiwe Zulu', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1), openedDate: new Date(2025, 11, 2), acknowledgedDate: new Date(2025, 11, 3), quizScore: 95, quizPassed: true },
  { id: 29, employeeId: 18, employeeName: 'Michelle September', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1), openedDate: new Date(2025, 11, 3), acknowledgedDate: new Date(2025, 11, 5), quizScore: 88, quizPassed: true },
  { id: 30, employeeId: 30, employeeName: 'Marike Joubert', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1), openedDate: new Date(2025, 11, 2), acknowledgedDate: new Date(2025, 11, 4), quizScore: 92, quizPassed: true },
  { id: 31, employeeId: 11, employeeName: 'Craig Williams', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'Sales', status: 'Overdue', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1), openedDate: new Date(2025, 11, 20) },
  { id: 32, employeeId: 17, employeeName: 'Bongani Sithole', policyId: 3, policyTitle: 'Employment Equity Policy', department: 'Sales', status: 'Overdue', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2026, 0, 1), sentDate: new Date(2025, 11, 1) },

  // Anti-Bribery - completed campaign high compliance
  { id: 33, employeeId: 2, employeeName: 'Pieter van der Merwe', policyId: 9, policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Executive', status: 'Acknowledged', assignedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1), sentDate: new Date(2025, 9, 1), openedDate: new Date(2025, 9, 1), acknowledgedDate: new Date(2025, 9, 2), quizScore: 85, quizPassed: true },
  { id: 34, employeeId: 5, employeeName: 'Johan Botha', policyId: 9, policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1), sentDate: new Date(2025, 9, 1), openedDate: new Date(2025, 9, 1), acknowledgedDate: new Date(2025, 9, 3), quizScore: 95, quizPassed: true },
  { id: 35, employeeId: 10, employeeName: 'Nomvula Khumalo', policyId: 9, policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Operations', status: 'Acknowledged', assignedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1), sentDate: new Date(2025, 9, 1), openedDate: new Date(2025, 9, 2), acknowledgedDate: new Date(2025, 9, 5), quizScore: 90, quizPassed: true },
  { id: 36, employeeId: 14, employeeName: 'Priya Govender', policyId: 9, policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Compliance', status: 'Acknowledged', assignedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1), sentDate: new Date(2025, 9, 1), openedDate: new Date(2025, 9, 1), acknowledgedDate: new Date(2025, 9, 2), quizScore: 100, quizPassed: true },
  { id: 37, employeeId: 21, employeeName: 'Lerato Mahlangu', policyId: 9, policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Procurement', status: 'Acknowledged', assignedDate: new Date(2025, 9, 1), dueDate: new Date(2025, 10, 1), sentDate: new Date(2025, 9, 1), openedDate: new Date(2025, 9, 3), acknowledgedDate: new Date(2025, 9, 6), quizScore: 88, quizPassed: true },

  // OHS Policy - mixed compliance
  { id: 38, employeeId: 38, employeeName: 'Grant Thompson', policyId: 4, policyTitle: 'Occupational Health & Safety Policy', department: 'Facilities', status: 'Acknowledged', assignedDate: new Date(2025, 8, 15), dueDate: new Date(2025, 9, 15), sentDate: new Date(2025, 8, 15), openedDate: new Date(2025, 8, 15), acknowledgedDate: new Date(2025, 8, 15), quizScore: 100, quizPassed: true },
  { id: 39, employeeId: 20, employeeName: 'Riyaad Jacobs', policyId: 4, policyTitle: 'Occupational Health & Safety Policy', department: 'Facilities', status: 'Acknowledged', assignedDate: new Date(2025, 8, 15), dueDate: new Date(2025, 9, 15), sentDate: new Date(2025, 8, 15), openedDate: new Date(2025, 8, 16), acknowledgedDate: new Date(2025, 8, 18), quizScore: 92, quizPassed: true },
  { id: 40, employeeId: 23, employeeName: 'Ayanda Zwane', policyId: 4, policyTitle: 'Occupational Health & Safety Policy', department: 'Operations', status: 'Acknowledged', assignedDate: new Date(2025, 8, 15), dueDate: new Date(2025, 9, 15), sentDate: new Date(2025, 8, 15), openedDate: new Date(2025, 8, 20), acknowledgedDate: new Date(2025, 8, 25), quizScore: 78, quizPassed: true },
  { id: 41, employeeId: 33, employeeName: 'Mandla Cele', policyId: 4, policyTitle: 'Occupational Health & Safety Policy', department: 'Operations', status: 'Acknowledged', assignedDate: new Date(2025, 8, 15), dueDate: new Date(2025, 9, 15), sentDate: new Date(2025, 8, 15), openedDate: new Date(2025, 8, 22), acknowledgedDate: new Date(2025, 9, 1), quizScore: 72, quizPassed: true },

  // King IV - Management governance
  { id: 42, employeeId: 2, employeeName: 'Pieter van der Merwe', policyId: 5, policyTitle: 'King IV Corporate Governance Policy', department: 'Executive', status: 'Acknowledged', assignedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 1, 15), sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 15), acknowledgedDate: new Date(2026, 0, 16) },
  { id: 43, employeeId: 4, employeeName: 'Fatima Patel', policyId: 5, policyTitle: 'King IV Corporate Governance Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 1, 15), sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 16), acknowledgedDate: new Date(2026, 0, 17) },
  { id: 44, employeeId: 5, employeeName: 'Johan Botha', policyId: 5, policyTitle: 'King IV Corporate Governance Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 1, 15), sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 15), acknowledgedDate: new Date(2026, 0, 16) },
  { id: 45, employeeId: 10, employeeName: 'Nomvula Khumalo', policyId: 5, policyTitle: 'King IV Corporate Governance Policy', department: 'Operations', status: 'Opened', assignedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 1, 15), sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 20) },
  { id: 46, employeeId: 11, employeeName: 'Craig Williams', policyId: 5, policyTitle: 'King IV Corporate Governance Policy', department: 'Sales', status: 'Overdue', assignedDate: new Date(2026, 0, 15), dueDate: new Date(2026, 0, 29), sentDate: new Date(2026, 0, 15) },

  // Expense & Travel - broad
  { id: 47, employeeId: 4, employeeName: 'Fatima Patel', policyId: 17, policyTitle: 'Expense & Travel Policy', department: 'Finance', status: 'Acknowledged', assignedDate: new Date(2025, 0, 5), dueDate: new Date(2025, 1, 5), sentDate: new Date(2025, 0, 5), openedDate: new Date(2025, 0, 5), acknowledgedDate: new Date(2025, 0, 6) },
  { id: 48, employeeId: 11, employeeName: 'Craig Williams', policyId: 17, policyTitle: 'Expense & Travel Policy', department: 'Sales', status: 'Acknowledged', assignedDate: new Date(2025, 0, 5), dueDate: new Date(2025, 1, 5), sentDate: new Date(2025, 0, 5), openedDate: new Date(2025, 0, 10), acknowledgedDate: new Date(2025, 0, 15) },
  { id: 49, employeeId: 17, employeeName: 'Bongani Sithole', policyId: 17, policyTitle: 'Expense & Travel Policy', department: 'Sales', status: 'Acknowledged', assignedDate: new Date(2025, 0, 5), dueDate: new Date(2025, 1, 5), sentDate: new Date(2025, 0, 5), openedDate: new Date(2025, 0, 12), acknowledgedDate: new Date(2025, 0, 20) },

  // Companies Act - Legal team
  { id: 50, employeeId: 5, employeeName: 'Johan Botha', policyId: 8, policyTitle: 'Companies Act Compliance Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2025, 2, 1), dueDate: new Date(2025, 3, 1), sentDate: new Date(2025, 2, 1), openedDate: new Date(2025, 2, 1), acknowledgedDate: new Date(2025, 2, 2) },
  { id: 51, employeeId: 19, employeeName: 'Vuyo Madonsela', policyId: 8, policyTitle: 'Companies Act Compliance Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2025, 2, 1), dueDate: new Date(2025, 3, 1), sentDate: new Date(2025, 2, 1), openedDate: new Date(2025, 2, 2), acknowledgedDate: new Date(2025, 2, 4) },
  { id: 52, employeeId: 29, employeeName: 'Imraan Moosa', policyId: 8, policyTitle: 'Companies Act Compliance Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2025, 2, 1), dueDate: new Date(2025, 3, 1), sentDate: new Date(2025, 2, 1), openedDate: new Date(2025, 2, 3), acknowledgedDate: new Date(2025, 2, 5) },
  { id: 53, employeeId: 40, employeeName: 'Anika Liebenberg', policyId: 8, policyTitle: 'Companies Act Compliance Policy', department: 'Legal', status: 'Acknowledged', assignedDate: new Date(2025, 2, 1), dueDate: new Date(2025, 3, 1), sentDate: new Date(2025, 2, 1), openedDate: new Date(2025, 2, 2), acknowledgedDate: new Date(2025, 2, 6) },

  // BBBEE - mixed
  { id: 54, employeeId: 3, employeeName: 'Naledi Mokoena', policyId: 2, policyTitle: 'BBBEE Compliance Policy', department: 'HR', status: 'Acknowledged', assignedDate: new Date(2026, 0, 13), dueDate: new Date(2026, 1, 28), sentDate: new Date(2026, 0, 13), openedDate: new Date(2026, 0, 13), acknowledgedDate: new Date(2026, 0, 14) },
  { id: 55, employeeId: 21, employeeName: 'Lerato Mahlangu', policyId: 2, policyTitle: 'BBBEE Compliance Policy', department: 'Procurement', status: 'Acknowledged', assignedDate: new Date(2026, 0, 13), dueDate: new Date(2026, 1, 28), sentDate: new Date(2026, 0, 13), openedDate: new Date(2026, 0, 14), acknowledgedDate: new Date(2026, 0, 16) },
  { id: 56, employeeId: 36, employeeName: 'Faizel Davids', policyId: 2, policyTitle: 'BBBEE Compliance Policy', department: 'Procurement', status: 'Opened', assignedDate: new Date(2026, 0, 13), dueDate: new Date(2026, 1, 28), sentDate: new Date(2026, 0, 13), openedDate: new Date(2026, 0, 20) },
  { id: 57, employeeId: 12, employeeName: 'Zanele Mthembu', policyId: 2, policyTitle: 'BBBEE Compliance Policy', department: 'Marketing', status: 'Pending', assignedDate: new Date(2026, 0, 13), dueDate: new Date(2026, 1, 28) },

  // Consumer Protection Act
  { id: 58, employeeId: 11, employeeName: 'Craig Williams', policyId: 16, policyTitle: 'Consumer Protection Act Compliance Policy', department: 'Sales', status: 'Acknowledged', assignedDate: new Date(2025, 5, 20), dueDate: new Date(2025, 6, 20), sentDate: new Date(2025, 5, 20), openedDate: new Date(2025, 5, 25), acknowledgedDate: new Date(2025, 6, 1), quizScore: 72, quizPassed: true },
  { id: 59, employeeId: 34, employeeName: 'Samantha Adams', policyId: 16, policyTitle: 'Consumer Protection Act Compliance Policy', department: 'Sales', status: 'Acknowledged', assignedDate: new Date(2025, 5, 20), dueDate: new Date(2025, 6, 20), sentDate: new Date(2025, 5, 20), openedDate: new Date(2025, 5, 22), acknowledgedDate: new Date(2025, 5, 28), quizScore: 85, quizPassed: true },

  // Whistleblower
  { id: 60, employeeId: 6, employeeName: 'Thandiwe Nkosi', policyId: 10, policyTitle: 'Whistleblower Protection Policy', department: 'Compliance', status: 'Acknowledged', assignedDate: new Date(2025, 9, 5), dueDate: new Date(2025, 10, 5), sentDate: new Date(2025, 9, 5), openedDate: new Date(2025, 9, 5), acknowledgedDate: new Date(2025, 9, 6) },
  { id: 61, employeeId: 14, employeeName: 'Priya Govender', policyId: 10, policyTitle: 'Whistleblower Protection Policy', department: 'Compliance', status: 'Acknowledged', assignedDate: new Date(2025, 9, 5), dueDate: new Date(2025, 10, 5), sentDate: new Date(2025, 9, 5), openedDate: new Date(2025, 9, 6), acknowledgedDate: new Date(2025, 9, 8) },

  // Remote Work - recent
  { id: 62, employeeId: 28, employeeName: 'Nozipho Buthelezi', policyId: 11, policyTitle: 'Remote Work & Flexible Arrangements Policy', department: 'Marketing', status: 'Acknowledged', assignedDate: new Date(2025, 10, 5), dueDate: new Date(2025, 11, 5), sentDate: new Date(2025, 10, 5), openedDate: new Date(2025, 10, 6), acknowledgedDate: new Date(2025, 10, 10) },
  { id: 63, employeeId: 39, employeeName: 'Zinhle Ngcobo', policyId: 11, policyTitle: 'Remote Work & Flexible Arrangements Policy', department: 'Marketing', status: 'Acknowledged', assignedDate: new Date(2025, 10, 5), dueDate: new Date(2025, 11, 5), sentDate: new Date(2025, 10, 5), openedDate: new Date(2025, 10, 8), acknowledgedDate: new Date(2025, 10, 12) },

  // Failed delivery
  { id: 64, employeeId: 33, employeeName: 'Mandla Cele', policyId: 1, policyTitle: 'POPIA Data Privacy Policy', department: 'Operations', status: 'Failed', assignedDate: new Date(2026, 0, 6), dueDate: new Date(2026, 1, 6), sentDate: new Date(2026, 0, 6) },

  // Procurement policy
  { id: 65, employeeId: 21, employeeName: 'Lerato Mahlangu', policyId: 13, policyTitle: 'Procurement & Supply Chain Policy', department: 'Procurement', status: 'Acknowledged', assignedDate: new Date(2025, 3, 20), dueDate: new Date(2025, 4, 20), sentDate: new Date(2025, 3, 20), openedDate: new Date(2025, 3, 20), acknowledgedDate: new Date(2025, 3, 21) },
];

// #endregion Acknowledgements

// #region Delegations

/** 10 delegation tasks for policy lifecycle management */
export const DEMO_DELEGATIONS: IDemoDelegation[] = [
  { id: 1, policyTitle: 'Social Media & Communications Policy', assignedTo: 'Zanele Mthembu', assignedBy: 'Thandiwe Nkosi', taskType: 'Draft', priority: 'Medium', status: 'InProgress', assignedDate: new Date(2026, 0, 5), dueDate: new Date(2026, 1, 5), notes: 'Draft new social media policy covering TikTok and emerging platforms. Align with POPIA requirements for employee data.' },
  { id: 2, policyTitle: 'Environmental, Social & Governance Policy', assignedTo: 'Priya Govender', assignedBy: 'Thandiwe Nkosi', taskType: 'Review', priority: 'High', status: 'InProgress', assignedDate: new Date(2026, 0, 10), dueDate: new Date(2026, 0, 31), notes: 'Review ESG policy v1.2 against latest JSE sustainability disclosure requirements and COP29 commitments.' },
  { id: 3, policyTitle: 'Disaster Recovery & Business Continuity Policy', assignedTo: 'Rajesh Naidoo', assignedBy: 'Sipho Dlamini', taskType: 'Review', priority: 'Critical', status: 'Overdue', assignedDate: new Date(2025, 11, 15), dueDate: new Date(2026, 0, 15), notes: 'Urgent review needed to incorporate updated load shedding stages and generator failover procedures for all offices.' },
  { id: 4, policyTitle: 'POPIA Data Privacy Policy', assignedTo: 'Dineo Masemola', assignedBy: 'Thandiwe Nkosi', taskType: 'Distribute', priority: 'High', status: 'Completed', assignedDate: new Date(2025, 11, 28), dueDate: new Date(2026, 0, 6), completedDate: new Date(2026, 0, 6), notes: 'Set up Q1 2026 POPIA awareness campaign for all employees. Ensure quiz is enabled.' },
  { id: 5, policyTitle: 'FICA & Anti-Money Laundering Policy', assignedTo: 'Karen Mostert', assignedBy: 'Fatima Patel', taskType: 'Review', priority: 'Medium', status: 'Completed', assignedDate: new Date(2025, 10, 1), dueDate: new Date(2025, 10, 30), completedDate: new Date(2025, 10, 22), notes: 'Annual review of FICA policy. Update CTR thresholds per latest FIC guidance note.' },
  { id: 6, policyTitle: 'Employment Equity Policy', assignedTo: 'Lindiwe Zulu', assignedBy: 'Naledi Mokoena', taskType: 'Distribute', priority: 'Medium', status: 'Pending', assignedDate: new Date(2026, 0, 25), dueDate: new Date(2026, 1, 10), notes: 'Distribute updated EE policy to Sales department. Craig Williams team has lowest compliance.' },
  { id: 7, policyTitle: 'King IV Corporate Governance Policy', assignedTo: 'Johan Botha', assignedBy: 'Pieter van der Merwe', taskType: 'Approve', priority: 'High', status: 'InProgress', assignedDate: new Date(2026, 0, 20), dueDate: new Date(2026, 1, 3), notes: 'Final approval required for King IV policy v1.9 incorporating updated board committee structures.' },
  { id: 8, policyTitle: 'Procurement & Supply Chain Policy', assignedTo: 'Faizel Davids', assignedBy: 'Lerato Mahlangu', taskType: 'Draft', priority: 'Low', status: 'Pending', assignedDate: new Date(2026, 0, 22), dueDate: new Date(2026, 2, 1), notes: 'Draft addendum for local supplier development programme aligned with BBBEE scorecard targets.' },
  { id: 9, policyTitle: 'Information Security Policy', assignedTo: 'Thabo Molefe', assignedBy: 'Sipho Dlamini', taskType: 'Review', priority: 'High', status: 'Completed', assignedDate: new Date(2025, 11, 1), dueDate: new Date(2025, 11, 20), completedDate: new Date(2025, 11, 18), notes: 'Technical review of InfoSec policy. Validate alignment with ISO 27001:2022 controls.' },
  { id: 10, policyTitle: 'Consumer Protection Act Compliance Policy', assignedTo: 'Vuyo Madonsela', assignedBy: 'Johan Botha', taskType: 'Review', priority: 'Medium', status: 'Pending', assignedDate: new Date(2026, 0, 28), dueDate: new Date(2026, 1, 28), notes: 'Review CPA policy against recent National Consumer Tribunal rulings affecting financial services.' },
];

// #endregion Delegations

// #region Team Compliance

/** Per-department team member compliance summary */
export const DEMO_TEAM_COMPLIANCE: IDemoTeamMember[] = [
  // Finance - highest compliance (98%)
  { id: 1, name: 'Fatima Patel', email: 'fatima.patel@proteafs.co.za', department: 'Finance', policiesAssigned: 8, policiesAcknowledged: 8, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 17) },
  { id: 2, name: 'Charl Pretorius', email: 'charl.pretorius@proteafs.co.za', department: 'Finance', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 14) },
  { id: 3, name: 'Annemarie du Plessis', email: 'annemarie.duplessis@proteafs.co.za', department: 'Finance', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 12) },
  { id: 4, name: 'Karen Mostert', email: 'karen.mostert@proteafs.co.za', department: 'Finance', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 16) },
  { id: 5, name: 'Yolande van Wyk', email: 'yolande.vanwyk@proteafs.co.za', department: 'Finance', policiesAssigned: 5, policiesAcknowledged: 5, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 10) },

  // Compliance - strong (95%)
  { id: 6, name: 'Thandiwe Nkosi', email: 'thandiwe.nkosi@proteafs.co.za', department: 'Compliance', policiesAssigned: 10, policiesAcknowledged: 10, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 20) },
  { id: 7, name: 'Priya Govender', email: 'priya.govender@proteafs.co.za', department: 'Compliance', policiesAssigned: 8, policiesAcknowledged: 8, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 18) },
  { id: 8, name: 'Ncumisa Dyani', email: 'ncumisa.dyani@proteafs.co.za', department: 'Compliance', policiesAssigned: 7, policiesAcknowledged: 6, policiesPending: 1, policiesOverdue: 0, compliancePercent: 86, lastActivity: new Date(2026, 0, 15) },
  { id: 9, name: 'Dineo Masemola', email: 'dineo.masemola@proteafs.co.za', department: 'Compliance', policiesAssigned: 7, policiesAcknowledged: 7, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 19) },

  // HR - solid (92%)
  { id: 10, name: 'Naledi Mokoena', email: 'naledi.mokoena@proteafs.co.za', department: 'HR', policiesAssigned: 9, policiesAcknowledged: 9, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 14) },
  { id: 11, name: 'Lindiwe Zulu', email: 'lindiwe.zulu@proteafs.co.za', department: 'HR', policiesAssigned: 6, policiesAcknowledged: 5, policiesPending: 1, policiesOverdue: 0, compliancePercent: 83, lastActivity: new Date(2026, 0, 10) },
  { id: 12, name: 'Michelle September', email: 'michelle.september@proteafs.co.za', department: 'HR', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 8) },
  { id: 13, name: 'Marike Joubert', email: 'marike.joubert@proteafs.co.za', department: 'HR', policiesAssigned: 6, policiesAcknowledged: 5, policiesPending: 1, policiesOverdue: 0, compliancePercent: 83, lastActivity: new Date(2026, 0, 12) },

  // IT - good (88%)
  { id: 14, name: 'Sipho Dlamini', email: 'sipho.dlamini@proteafs.co.za', department: 'IT', policiesAssigned: 8, policiesAcknowledged: 8, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 9) },
  { id: 15, name: 'Rajesh Naidoo', email: 'rajesh.naidoo@proteafs.co.za', department: 'IT', policiesAssigned: 7, policiesAcknowledged: 6, policiesPending: 1, policiesOverdue: 0, compliancePercent: 86, lastActivity: new Date(2026, 0, 10) },
  { id: 16, name: 'Thabo Molefe', email: 'thabo.molefe@proteafs.co.za', department: 'IT', policiesAssigned: 6, policiesAcknowledged: 5, policiesPending: 1, policiesOverdue: 0, compliancePercent: 83, lastActivity: new Date(2026, 0, 12) },
  { id: 17, name: 'Suresh Pillay', email: 'suresh.pillay@proteafs.co.za', department: 'IT', policiesAssigned: 6, policiesAcknowledged: 4, policiesPending: 2, policiesOverdue: 0, compliancePercent: 67, lastActivity: new Date(2026, 0, 15) },
  { id: 18, name: 'Sibusiso Ndlovu', email: 'sibusiso.ndlovu@proteafs.co.za', department: 'IT', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 11) },
  { id: 19, name: 'Charlize Steyn', email: 'charlize.steyn@proteafs.co.za', department: 'IT', policiesAssigned: 6, policiesAcknowledged: 4, policiesPending: 2, policiesOverdue: 0, compliancePercent: 67, lastActivity: new Date(2026, 0, 8) },

  // Sales - struggling (72%)
  { id: 20, name: 'Craig Williams', email: 'craig.williams@proteafs.co.za', department: 'Sales', policiesAssigned: 7, policiesAcknowledged: 3, policiesPending: 2, policiesOverdue: 2, compliancePercent: 43, lastActivity: new Date(2026, 0, 15) },
  { id: 21, name: 'Bongani Sithole', email: 'bongani.sithole@proteafs.co.za', department: 'Sales', policiesAssigned: 6, policiesAcknowledged: 2, policiesPending: 2, policiesOverdue: 2, compliancePercent: 33, lastActivity: new Date(2025, 11, 1) },
  { id: 22, name: 'Tshepo Mabaso', email: 'tshepo.mabaso@proteafs.co.za', department: 'Sales', policiesAssigned: 5, policiesAcknowledged: 2, policiesPending: 2, policiesOverdue: 1, compliancePercent: 40, lastActivity: new Date(2026, 0, 12) },
  { id: 23, name: 'Samantha Adams', email: 'samantha.adams@proteafs.co.za', department: 'Sales', policiesAssigned: 5, policiesAcknowledged: 4, policiesPending: 1, policiesOverdue: 0, compliancePercent: 80, lastActivity: new Date(2026, 0, 6) },
  { id: 24, name: 'Lungelo Mkhize', email: 'lungelo.mkhize@proteafs.co.za', department: 'Sales', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 10) },

  // Legal - strong (94%)
  { id: 25, name: 'Johan Botha', email: 'johan.botha@proteafs.co.za', department: 'Legal', policiesAssigned: 9, policiesAcknowledged: 9, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 16) },
  { id: 26, name: 'Vuyo Madonsela', email: 'vuyo.madonsela@proteafs.co.za', department: 'Legal', policiesAssigned: 7, policiesAcknowledged: 7, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 14) },
  { id: 27, name: 'Imraan Moosa', email: 'imraan.moosa@proteafs.co.za', department: 'Legal', policiesAssigned: 6, policiesAcknowledged: 5, policiesPending: 1, policiesOverdue: 0, compliancePercent: 83, lastActivity: new Date(2026, 0, 10) },
  { id: 28, name: 'Anika Liebenberg', email: 'anika.liebenberg@proteafs.co.za', department: 'Legal', policiesAssigned: 5, policiesAcknowledged: 5, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 8) },

  // Operations - moderate (85%)
  { id: 29, name: 'Nomvula Khumalo', email: 'nomvula.khumalo@proteafs.co.za', department: 'Operations', policiesAssigned: 7, policiesAcknowledged: 6, policiesPending: 1, policiesOverdue: 0, compliancePercent: 86, lastActivity: new Date(2026, 0, 20) },
  { id: 30, name: 'Ayanda Zwane', email: 'ayanda.zwane@proteafs.co.za', department: 'Operations', policiesAssigned: 5, policiesAcknowledged: 4, policiesPending: 1, policiesOverdue: 0, compliancePercent: 80, lastActivity: new Date(2026, 0, 5) },
  { id: 31, name: 'Mandla Cele', email: 'mandla.cele@proteafs.co.za', department: 'Operations', policiesAssigned: 5, policiesAcknowledged: 4, policiesPending: 0, policiesOverdue: 1, compliancePercent: 80, lastActivity: new Date(2025, 9, 1) },

  // Marketing (82%)
  { id: 32, name: 'Zanele Mthembu', email: 'zanele.mthembu@proteafs.co.za', department: 'Marketing', policiesAssigned: 5, policiesAcknowledged: 3, policiesPending: 2, policiesOverdue: 0, compliancePercent: 60, lastActivity: new Date(2026, 0, 6) },
  { id: 33, name: 'Nozipho Buthelezi', email: 'nozipho.buthelezi@proteafs.co.za', department: 'Marketing', policiesAssigned: 4, policiesAcknowledged: 4, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2025, 10, 10) },
  { id: 34, name: 'Zinhle Ngcobo', email: 'zinhle.ngcobo@proteafs.co.za', department: 'Marketing', policiesAssigned: 4, policiesAcknowledged: 3, policiesPending: 1, policiesOverdue: 0, compliancePercent: 75, lastActivity: new Date(2025, 10, 12) },

  // Engineering (86%)
  { id: 35, name: 'Hendrik Viljoen', email: 'hendrik.viljoen@proteafs.co.za', department: 'Engineering', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 9) },
  { id: 36, name: 'Willem Erasmus', email: 'willem.erasmus@proteafs.co.za', department: 'Engineering', policiesAssigned: 5, policiesAcknowledged: 3, policiesPending: 1, policiesOverdue: 1, compliancePercent: 60, lastActivity: new Date(2026, 0, 14) },
  { id: 37, name: 'Kamohelo Tlali', email: 'kamohelo.tlali@proteafs.co.za', department: 'Engineering', policiesAssigned: 5, policiesAcknowledged: 5, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 13) },

  // Remaining departments
  { id: 38, name: 'Riyaad Jacobs', email: 'riyaad.jacobs@proteafs.co.za', department: 'Facilities', policiesAssigned: 5, policiesAcknowledged: 5, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2025, 8, 18) },
  { id: 39, name: 'Grant Thompson', email: 'grant.thompson@proteafs.co.za', department: 'Facilities', policiesAssigned: 5, policiesAcknowledged: 5, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2025, 8, 15) },
  { id: 40, name: 'Lerato Mahlangu', email: 'lerato.mahlangu@proteafs.co.za', department: 'Procurement', policiesAssigned: 6, policiesAcknowledged: 6, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 16) },
  { id: 41, name: 'Faizel Davids', email: 'faizel.davids@proteafs.co.za', department: 'Procurement', policiesAssigned: 5, policiesAcknowledged: 3, policiesPending: 2, policiesOverdue: 0, compliancePercent: 60, lastActivity: new Date(2026, 0, 20) },
  { id: 42, name: 'Pieter van der Merwe', email: 'pieter.vandermerwe@proteafs.co.za', department: 'Executive', policiesAssigned: 10, policiesAcknowledged: 10, policiesPending: 0, policiesOverdue: 0, compliancePercent: 100, lastActivity: new Date(2026, 0, 16) },
];

// #endregion Team Compliance

// #region SLA Metrics

/** SLA tracking for policy lifecycle management */
export const DEMO_SLA_METRICS: IDemoSlaMetric[] = [
  {
    type: 'Policy Review Cycle', targetDays: 30, actualAvgDays: 26, percentMet: 88, status: 'Met',
    breaches: [
      { policyTitle: 'Disaster Recovery & Business Continuity Policy', targetDays: 30, actualDays: 47, breachDate: new Date(2026, 0, 15), department: 'IT' },
    ]
  },
  {
    type: 'Policy Acknowledgement', targetDays: 14, actualAvgDays: 11, percentMet: 82, status: 'At Risk',
    breaches: [
      { policyTitle: 'POPIA Data Privacy Policy', targetDays: 14, actualDays: 25, breachDate: new Date(2026, 0, 31), department: 'Sales' },
      { policyTitle: 'Employment Equity Policy', targetDays: 14, actualDays: 60, breachDate: new Date(2026, 0, 30), department: 'Sales' },
      { policyTitle: 'Information Security Policy', targetDays: 14, actualDays: 23, breachDate: new Date(2026, 0, 31), department: 'Engineering' },
    ]
  },
  {
    type: 'Policy Approval', targetDays: 7, actualAvgDays: 5, percentMet: 94, status: 'Met',
    breaches: [
      { policyTitle: 'Environmental, Social & Governance Policy', targetDays: 7, actualDays: 12, breachDate: new Date(2026, 0, 27), department: 'Compliance' },
    ]
  },
  {
    type: 'Delegation Completion', targetDays: 21, actualAvgDays: 18, percentMet: 80, status: 'At Risk',
    breaches: [
      { policyTitle: 'Disaster Recovery & Business Continuity Policy', targetDays: 21, actualDays: 47, breachDate: new Date(2026, 0, 31), department: 'IT' },
      { policyTitle: 'Social Media & Communications Policy', targetDays: 21, actualDays: 26, breachDate: new Date(2026, 0, 31), department: 'Marketing' },
    ]
  },
  {
    type: 'Incident Response', targetDays: 3, actualAvgDays: 2, percentMet: 92, status: 'Met',
    breaches: [
      { policyTitle: 'Information Security Policy', targetDays: 3, actualDays: 5, breachDate: new Date(2025, 11, 18), department: 'IT' },
    ]
  },
  {
    type: 'Campaign Distribution', targetDays: 5, actualAvgDays: 3, percentMet: 96, status: 'Met',
    breaches: []
  },
  {
    type: 'Violation Resolution', targetDays: 14, actualAvgDays: 16, percentMet: 70, status: 'Breached',
    breaches: [
      { policyTitle: 'POPIA Data Privacy Policy', targetDays: 14, actualDays: 22, breachDate: new Date(2025, 11, 10), department: 'Sales' },
      { policyTitle: 'Anti-Bribery & Corruption Policy', targetDays: 14, actualDays: 19, breachDate: new Date(2025, 10, 25), department: 'Procurement' },
      { policyTitle: 'Information Security Policy', targetDays: 14, actualDays: 28, breachDate: new Date(2025, 11, 28), department: 'Operations' },
    ]
  },
];

// #endregion SLA Metrics

// #region Audit Log

/** 35 audit log entries covering various policy management activities */
export const DEMO_AUDIT_LOG: IDemoAuditEntry[] = [
  { id: 1, timestamp: new Date(2026, 0, 30, 14, 22), user: 'Thandiwe Nkosi', action: 'Policy Updated', policyTitle: 'Environmental, Social & Governance Policy', details: 'Updated ESG policy v1.2 with new JSE sustainability disclosure requirements.', ipAddress: '196.25.47.102' },
  { id: 2, timestamp: new Date(2026, 0, 29, 9, 15), user: 'Sipho Dlamini', action: 'Campaign Created', policyTitle: 'POPIA Data Privacy Policy', details: 'Created scheduled OHS Policy Update Distribution campaign for February 2026.', ipAddress: '196.25.47.55' },
  { id: 3, timestamp: new Date(2026, 0, 28, 16, 45), user: 'Naledi Mokoena', action: 'Campaign Created', policyTitle: 'Employment Equity Policy', details: 'Created draft campaign for Employment Equity Awareness targeting Sales department.', ipAddress: '196.25.47.88' },
  { id: 4, timestamp: new Date(2026, 0, 27, 11, 30), user: 'Thandiwe Nkosi', action: 'Delegation Created', policyTitle: 'Consumer Protection Act Compliance Policy', details: 'Delegated CPA policy review to Vuyo Madonsela in Legal.', ipAddress: '196.25.47.102' },
  { id: 5, timestamp: new Date(2026, 0, 25, 10, 0), user: 'Naledi Mokoena', action: 'Delegation Created', policyTitle: 'Employment Equity Policy', details: 'Delegated EE policy distribution to Lindiwe Zulu for Sales department rollout.', ipAddress: '196.25.47.88' },
  { id: 6, timestamp: new Date(2026, 0, 22, 14, 10), user: 'Lerato Mahlangu', action: 'Delegation Created', policyTitle: 'Procurement & Supply Chain Policy', details: 'Assigned procurement policy addendum draft to Faizel Davids.', ipAddress: '196.25.47.120' },
  { id: 7, timestamp: new Date(2026, 0, 20, 8, 45), user: 'Pieter van der Merwe', action: 'Delegation Created', policyTitle: 'King IV Corporate Governance Policy', details: 'Assigned final approval of King IV policy v1.9 to Johan Botha.', ipAddress: '196.25.47.10' },
  { id: 8, timestamp: new Date(2026, 0, 20, 15, 30), user: 'Thandiwe Nkosi', action: 'Campaign Paused', policyTitle: 'Environmental, Social & Governance Policy', details: 'Paused ESG policy review campaign pending updated JSE guidelines.', ipAddress: '196.25.47.102' },
  { id: 9, timestamp: new Date(2026, 0, 17, 9, 0), user: 'Fatima Patel', action: 'Policy Acknowledged', policyTitle: 'King IV Corporate Governance Policy', details: 'Acknowledged King IV Corporate Governance Policy v1.8.', ipAddress: '41.13.252.78' },
  { id: 10, timestamp: new Date(2026, 0, 16, 11, 20), user: 'Pieter van der Merwe', action: 'Policy Acknowledged', policyTitle: 'King IV Corporate Governance Policy', details: 'Acknowledged King IV Corporate Governance Policy v1.8.', ipAddress: '196.25.47.10' },
  { id: 11, timestamp: new Date(2026, 0, 15, 8, 0), user: 'Johan Botha', action: 'Campaign Launched', policyTitle: 'Management Governance Pack', details: 'Launched Management Governance Pack 2026 campaign targeting executive and management teams.', ipAddress: '41.72.128.15' },
  { id: 12, timestamp: new Date(2026, 0, 13, 9, 30), user: 'Thandiwe Nkosi', action: 'Campaign Launched', policyTitle: 'Annual Regulatory Compliance Pack', details: 'Launched Annual Regulatory Compliance 2026 campaign for all employees.', ipAddress: '196.25.47.102' },
  { id: 13, timestamp: new Date(2026, 0, 10, 14, 15), user: 'Thandiwe Nkosi', action: 'Delegation Created', policyTitle: 'Environmental, Social & Governance Policy', details: 'Delegated ESG policy review to Priya Govender.', ipAddress: '196.25.47.102' },
  { id: 14, timestamp: new Date(2026, 0, 10, 10, 0), user: 'Charlize Steyn', action: 'Policy Acknowledged', policyTitle: 'New Employee Onboarding Pack', details: 'Completed onboarding pack acknowledgement (6 policies).', ipAddress: '41.185.22.44' },
  { id: 15, timestamp: new Date(2026, 0, 9, 16, 30), user: 'Sipho Dlamini', action: 'Policy Acknowledged', policyTitle: 'Information Security Policy', details: 'Acknowledged Information Security Policy v5.2. Quiz score: 100%.', ipAddress: '196.25.47.55' },
  { id: 16, timestamp: new Date(2026, 0, 8, 8, 30), user: 'Sipho Dlamini', action: 'Campaign Launched', policyTitle: 'IT Security Essentials Pack', details: 'Launched IT Security Policy Rollout campaign for IT and Engineering departments.', ipAddress: '196.25.47.55' },
  { id: 17, timestamp: new Date(2026, 0, 7, 11, 45), user: 'Sipho Dlamini', action: 'Quiz Score Recorded', policyTitle: 'POPIA Data Privacy Policy', details: 'Sipho Dlamini scored 95% on POPIA quiz (passed).', ipAddress: '196.25.47.55' },
  { id: 18, timestamp: new Date(2026, 0, 6, 9, 0), user: 'Dineo Masemola', action: 'Campaign Launched', policyTitle: 'POPIA Data Privacy Policy', details: 'Launched Q1 2026 POPIA Awareness Campaign for all employees.', ipAddress: '196.25.47.95' },
  { id: 19, timestamp: new Date(2026, 0, 6, 8, 0), user: 'Thandiwe Nkosi', action: 'Delegation Completed', policyTitle: 'POPIA Data Privacy Policy', details: 'Dineo Masemola completed POPIA campaign distribution delegation.', ipAddress: '196.25.47.102' },
  { id: 20, timestamp: new Date(2026, 0, 5, 10, 0), user: 'Thandiwe Nkosi', action: 'Delegation Created', policyTitle: 'Social Media & Communications Policy', details: 'Delegated social media policy draft to Zanele Mthembu.', ipAddress: '196.25.47.102' },
  { id: 21, timestamp: new Date(2025, 11, 28, 15, 0), user: 'Thandiwe Nkosi', action: 'Campaign Created', policyTitle: 'Annual Regulatory Compliance Pack', details: 'Created Annual Regulatory Compliance 2026 campaign.', ipAddress: '196.25.47.102' },
  { id: 22, timestamp: new Date(2025, 11, 20, 9, 30), user: 'Thandiwe Nkosi', action: 'Campaign Created', policyTitle: 'POPIA Data Privacy Policy', details: 'Created Q1 2026 POPIA Awareness Campaign.', ipAddress: '196.25.47.102' },
  { id: 23, timestamp: new Date(2025, 11, 18, 14, 0), user: 'Thabo Molefe', action: 'Delegation Completed', policyTitle: 'Information Security Policy', details: 'Completed technical review of InfoSec policy. Confirmed ISO 27001:2022 alignment.', ipAddress: '196.25.47.62' },
  { id: 24, timestamp: new Date(2025, 11, 15, 11, 30), user: 'Naledi Mokoena', action: 'Campaign Launched', policyTitle: 'New Employee Onboarding Pack', details: 'Launched onboarding campaign for January 2026 new joiners.', ipAddress: '196.25.47.88' },
  { id: 25, timestamp: new Date(2025, 11, 10, 16, 0), user: 'Thandiwe Nkosi', action: 'Violation Reported', policyTitle: 'POPIA Data Privacy Policy', details: 'Reported POPIA breach in Sales department: customer data shared via personal email.', ipAddress: '196.25.47.102' },
  { id: 26, timestamp: new Date(2025, 11, 1, 9, 0), user: 'Sipho Dlamini', action: 'Delegation Created', policyTitle: 'Disaster Recovery & Business Continuity Policy', details: 'Delegated DR policy review to Rajesh Naidoo. Critical priority.', ipAddress: '196.25.47.55' },
  { id: 27, timestamp: new Date(2025, 10, 25, 14, 30), user: 'Fatima Patel', action: 'Campaign Completed', policyTitle: 'Finance & Risk Pack', details: 'Finance Team FICA Refresher campaign completed. 100% compliance achieved.', ipAddress: '41.13.252.78' },
  { id: 28, timestamp: new Date(2025, 10, 22, 10, 15), user: 'Karen Mostert', action: 'Delegation Completed', policyTitle: 'FICA & Anti-Money Laundering Policy', details: 'Completed annual FICA policy review. Updated CTR thresholds.', ipAddress: '196.25.47.130' },
  { id: 29, timestamp: new Date(2025, 10, 5, 9, 0), user: 'Thandiwe Nkosi', action: 'Violation Resolved', policyTitle: 'Anti-Bribery & Corruption Policy', details: 'Resolved investigation into unauthorised gifts to Procurement staff. Disciplinary action taken.', ipAddress: '196.25.47.102' },
  { id: 30, timestamp: new Date(2025, 9, 28, 16, 45), user: 'Thandiwe Nkosi', action: 'Campaign Completed', policyTitle: 'Anti-Bribery & Corruption Policy', details: 'Anti-Bribery Annual Refresher campaign completed. 93% compliance rate.', ipAddress: '196.25.47.102' },
  { id: 31, timestamp: new Date(2025, 9, 15, 11, 0), user: 'Grant Thompson', action: 'Policy Published', policyTitle: 'Occupational Health & Safety Policy', details: 'Published OHS Policy v4.0 with updated emergency evacuation procedures.', ipAddress: '196.25.47.145' },
  { id: 32, timestamp: new Date(2025, 9, 1, 8, 0), user: 'Thandiwe Nkosi', action: 'Campaign Launched', policyTitle: 'Anti-Bribery & Corruption Policy', details: 'Launched Anti-Bribery Annual Refresher campaign for all employees.', ipAddress: '196.25.47.102' },
  { id: 33, timestamp: new Date(2025, 8, 20, 14, 30), user: 'Sipho Dlamini', action: 'Policy Updated', policyTitle: 'Information Security Policy', details: 'Updated InfoSec Policy to v5.2. Added AI usage guidelines and deepfake awareness section.', ipAddress: '196.25.47.55' },
  { id: 34, timestamp: new Date(2025, 8, 1, 9, 0), user: 'Naledi Mokoena', action: 'Policy Published', policyTitle: 'Remote Work & Flexible Arrangements Policy', details: 'Published Remote Work Policy v2.0 with updated hybrid working guidelines.', ipAddress: '196.25.47.88' },
  { id: 35, timestamp: new Date(2025, 7, 15, 10, 30), user: 'Johan Botha', action: 'Policy Archived', policyTitle: 'Code of Ethics & Business Conduct', details: 'Archived Code of Ethics v2.0. Superseded by updated Anti-Bribery and Whistleblower policies.', ipAddress: '41.72.128.15' },
];

// #endregion Audit Log

// #region Violations

/** 12 policy violations at various severity levels and resolution statuses */
export const DEMO_VIOLATIONS: IDemoViolation[] = [
  { id: 1, severity: 'Critical', policyTitle: 'POPIA Data Privacy Policy', department: 'Sales', description: 'Customer personal information (ID numbers and bank details) shared via unencrypted personal email to third-party vendor without data subject consent.', status: 'Under Investigation', reportedDate: new Date(2025, 11, 10) },
  { id: 2, severity: 'High', policyTitle: 'Anti-Bribery & Corruption Policy', department: 'Procurement', description: 'Unauthorised gifts valued at R15,000 accepted from prospective supplier during tender process.', status: 'Resolved', reportedDate: new Date(2025, 9, 8), resolvedDate: new Date(2025, 10, 5) },
  { id: 3, severity: 'High', policyTitle: 'Information Security Policy', department: 'Operations', description: 'Shared login credentials for core banking system found posted on shared whiteboard in Durban office.', status: 'Resolved', reportedDate: new Date(2025, 10, 30), resolvedDate: new Date(2025, 11, 28) },
  { id: 4, severity: 'Medium', policyTitle: 'FICA & Anti-Money Laundering Policy', department: 'Finance', description: 'Client onboarding completed without full FICA documentation. Missing proof of address for three high-value clients.', status: 'Resolved', reportedDate: new Date(2025, 8, 15), resolvedDate: new Date(2025, 8, 28) },
  { id: 5, severity: 'Medium', policyTitle: 'Employment Equity Policy', department: 'Sales', description: 'Recruitment shortlist for Senior Sales Executive position did not include any candidates from designated groups, contrary to EE plan targets.', status: 'Open', reportedDate: new Date(2026, 0, 15) },
  { id: 6, severity: 'Low', policyTitle: 'Expense & Travel Policy', department: 'Marketing', description: 'Expense claim submitted 45 days after travel date, exceeding 14-day submission policy. Amount: R8,200.', status: 'Resolved', reportedDate: new Date(2025, 11, 5), resolvedDate: new Date(2025, 11, 12) },
  { id: 7, severity: 'High', policyTitle: 'POPIA Data Privacy Policy', department: 'IT', description: 'Development database containing production customer data discovered accessible without authentication on internal network.', status: 'Resolved', reportedDate: new Date(2025, 7, 20), resolvedDate: new Date(2025, 7, 22) },
  { id: 8, severity: 'Medium', policyTitle: 'Occupational Health & Safety Policy', department: 'Facilities', description: 'Fire extinguisher inspection overdue by 3 months in Johannesburg head office. Multiple units found past certification date.', status: 'Resolved', reportedDate: new Date(2025, 10, 1), resolvedDate: new Date(2025, 10, 8) },
  { id: 9, severity: 'Critical', policyTitle: 'Information Security Policy', department: 'IT', description: 'Suspected ransomware phishing email opened by employee. Endpoint isolated. Forensic investigation in progress.', status: 'Escalated', reportedDate: new Date(2026, 0, 22) },
  { id: 10, severity: 'Low', policyTitle: 'Remote Work & Flexible Arrangements Policy', department: 'Engineering', description: 'Employee working remotely from undisclosed location outside South Africa without prior approval for 2 weeks.', status: 'Open', reportedDate: new Date(2026, 0, 18) },
  { id: 11, severity: 'Medium', policyTitle: 'BBBEE Compliance Policy', department: 'Procurement', description: 'Q4 2025 procurement spend analysis shows only 38% to BBBEE Level 1-4 suppliers against 60% target.', status: 'Under Investigation', reportedDate: new Date(2026, 0, 10) },
  { id: 12, severity: 'Low', policyTitle: 'King IV Corporate Governance Policy', department: 'Executive', description: 'Board meeting minutes not circulated within required 5-day window. Distributed 8 days after meeting.', status: 'Resolved', reportedDate: new Date(2025, 11, 20), resolvedDate: new Date(2025, 11, 22) },
];

// #endregion Violations

// #region Quiz Results

/** 25 quiz attempt records showing pass/fail patterns */
export const DEMO_QUIZ_RESULTS: IDemoQuizResult[] = [
  { id: 1, employeeName: 'Sipho Dlamini', department: 'IT', policyTitle: 'POPIA Data Privacy Policy', score: 95, passed: true, attemptDate: new Date(2026, 0, 7), attemptNumber: 1, timeTaken: 12 },
  { id: 2, employeeName: 'Naledi Mokoena', department: 'HR', policyTitle: 'POPIA Data Privacy Policy', score: 90, passed: true, attemptDate: new Date(2026, 0, 8), attemptNumber: 1, timeTaken: 15 },
  { id: 3, employeeName: 'Fatima Patel', department: 'Finance', policyTitle: 'POPIA Data Privacy Policy', score: 100, passed: true, attemptDate: new Date(2026, 0, 9), attemptNumber: 1, timeTaken: 8 },
  { id: 4, employeeName: 'Thandiwe Nkosi', department: 'Compliance', policyTitle: 'POPIA Data Privacy Policy', score: 100, passed: true, attemptDate: new Date(2026, 0, 6), attemptNumber: 1, timeTaken: 7 },
  { id: 5, employeeName: 'Rajesh Naidoo', department: 'IT', policyTitle: 'POPIA Data Privacy Policy', score: 85, passed: true, attemptDate: new Date(2026, 0, 10), attemptNumber: 1, timeTaken: 18 },
  { id: 6, employeeName: 'Charl Pretorius', department: 'Finance', policyTitle: 'POPIA Data Privacy Policy', score: 90, passed: true, attemptDate: new Date(2026, 0, 8), attemptNumber: 1, timeTaken: 14 },
  { id: 7, employeeName: 'Fatima Patel', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 100, passed: true, attemptDate: new Date(2025, 10, 2), attemptNumber: 1, timeTaken: 10 },
  { id: 8, employeeName: 'Charl Pretorius', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 95, passed: true, attemptDate: new Date(2025, 10, 3), attemptNumber: 1, timeTaken: 16 },
  { id: 9, employeeName: 'Annemarie du Plessis', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 88, passed: true, attemptDate: new Date(2025, 10, 4), attemptNumber: 1, timeTaken: 20 },
  { id: 10, employeeName: 'Karen Mostert', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 92, passed: true, attemptDate: new Date(2025, 10, 5), attemptNumber: 1, timeTaken: 13 },
  { id: 11, employeeName: 'Yolande van Wyk', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 60, passed: false, attemptDate: new Date(2025, 10, 3), attemptNumber: 1, timeTaken: 22 },
  { id: 12, employeeName: 'Yolande van Wyk', department: 'Finance', policyTitle: 'FICA & Anti-Money Laundering Policy', score: 85, passed: true, attemptDate: new Date(2025, 10, 5), attemptNumber: 2, timeTaken: 18 },
  { id: 13, employeeName: 'Sipho Dlamini', department: 'IT', policyTitle: 'Information Security Policy', score: 100, passed: true, attemptDate: new Date(2026, 0, 9), attemptNumber: 1, timeTaken: 9 },
  { id: 14, employeeName: 'Rajesh Naidoo', department: 'IT', policyTitle: 'Information Security Policy', score: 92, passed: true, attemptDate: new Date(2026, 0, 10), attemptNumber: 1, timeTaken: 14 },
  { id: 15, employeeName: 'Thabo Molefe', department: 'IT', policyTitle: 'Information Security Policy', score: 88, passed: true, attemptDate: new Date(2026, 0, 12), attemptNumber: 1, timeTaken: 17 },
  { id: 16, employeeName: 'Sibusiso Ndlovu', department: 'IT', policyTitle: 'Information Security Policy', score: 95, passed: true, attemptDate: new Date(2026, 0, 11), attemptNumber: 1, timeTaken: 11 },
  { id: 17, employeeName: 'Hendrik Viljoen', department: 'Engineering', policyTitle: 'Information Security Policy', score: 90, passed: true, attemptDate: new Date(2026, 0, 9), attemptNumber: 1, timeTaken: 13 },
  { id: 18, employeeName: 'Kamohelo Tlali', department: 'Engineering', policyTitle: 'Information Security Policy', score: 55, passed: false, attemptDate: new Date(2026, 0, 11), attemptNumber: 1, timeTaken: 25 },
  { id: 19, employeeName: 'Kamohelo Tlali', department: 'Engineering', policyTitle: 'Information Security Policy', score: 80, passed: true, attemptDate: new Date(2026, 0, 13), attemptNumber: 2, timeTaken: 19 },
  { id: 20, employeeName: 'Craig Williams', department: 'Sales', policyTitle: 'Consumer Protection Act Compliance Policy', score: 72, passed: true, attemptDate: new Date(2025, 6, 1), attemptNumber: 1, timeTaken: 20 },
  { id: 21, employeeName: 'Samantha Adams', department: 'Sales', policyTitle: 'Consumer Protection Act Compliance Policy', score: 85, passed: true, attemptDate: new Date(2025, 5, 28), attemptNumber: 1, timeTaken: 16 },
  { id: 22, employeeName: 'Pieter van der Merwe', department: 'Executive', policyTitle: 'Anti-Bribery & Corruption Policy', score: 85, passed: true, attemptDate: new Date(2025, 9, 2), attemptNumber: 1, timeTaken: 15 },
  { id: 23, employeeName: 'Priya Govender', department: 'Compliance', policyTitle: 'Anti-Bribery & Corruption Policy', score: 100, passed: true, attemptDate: new Date(2025, 9, 2), attemptNumber: 1, timeTaken: 8 },
  { id: 24, employeeName: 'Grant Thompson', department: 'Facilities', policyTitle: 'Occupational Health & Safety Policy', score: 100, passed: true, attemptDate: new Date(2025, 8, 15), attemptNumber: 1, timeTaken: 10 },
  { id: 25, employeeName: 'Ayanda Zwane', department: 'Operations', policyTitle: 'Occupational Health & Safety Policy', score: 78, passed: true, attemptDate: new Date(2025, 8, 25), attemptNumber: 1, timeTaken: 22 },
];

// #endregion Quiz Results

// #region Helper Function

/**
 * Returns a summary of all demo data for dashboard overview displays.
 * Computes totals and rates from the underlying data arrays.
 */
export function getDemoDataSummary(): IDemoDataSummary {
  const totalEmployees = DEMO_EMPLOYEES.length;
  const totalPolicies = DEMO_POLICIES.length;
  const publishedPolicies = DEMO_POLICIES.filter(p => p.status === 'Published').length;
  const draftPolicies = DEMO_POLICIES.filter(p => p.status === 'Draft').length;

  const totalAcknowledgements = DEMO_ACKNOWLEDGEMENTS.length;
  const acknowledgedCount = DEMO_ACKNOWLEDGEMENTS.filter(a => a.status === 'Acknowledged').length;
  const pendingCount = DEMO_ACKNOWLEDGEMENTS.filter(a => ['Pending', 'Sent', 'Delivered', 'Opened'].includes(a.status)).length;
  const overdueCount = DEMO_ACKNOWLEDGEMENTS.filter(a => a.status === 'Overdue').length;
  const overallComplianceRate = Math.round((acknowledgedCount / totalAcknowledgements) * 100);

  const activeCampaigns = DEMO_CAMPAIGNS.filter(c => c.status === 'Active').length;
  const completedCampaigns = DEMO_CAMPAIGNS.filter(c => c.status === 'Completed').length;

  const openViolations = DEMO_VIOLATIONS.filter(v => v.status === 'Open' || v.status === 'Under Investigation' || v.status === 'Escalated').length;

  const quizScores = DEMO_QUIZ_RESULTS.map(q => q.score);
  const averageQuizScore = Math.round(quizScores.reduce((sum, s) => sum + s, 0) / quizScores.length);

  const departments = new Set(DEMO_EMPLOYEES.map(e => e.department));
  const departmentsTracked = departments.size;

  return {
    totalEmployees,
    totalPolicies,
    publishedPolicies,
    draftPolicies,
    overallComplianceRate,
    overdueCount,
    activeCampaigns,
    completedCampaigns,
    totalAcknowledgements,
    acknowledgedCount,
    pendingCount,
    openViolations,
    averageQuizScore,
    departmentsTracked,
  };
}

// #endregion Helper Function
