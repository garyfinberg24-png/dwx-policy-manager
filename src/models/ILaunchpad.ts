import { UserRole } from '../services/RoleDetectionService';

export interface ILaunchpadTile {
  id: string;
  title: string;
  description: string;
  icon: string;
  iconColor: string;
  category: LaunchpadCategory;
  url: string;
  isExternal?: boolean;
  permissions?: string[];
  allowedRoles?: UserRole[]; // Role-based access control
  badge?: string;
  badgeColor?: string;
}

export enum LaunchpadCategory {
  ProcessManagement = 'Process Management',
  TaskManagement = 'Task Management',
  TalentManagement = 'Talent Management',
  AssetManagement = 'Asset Management',
  EmployeeEngagement = 'Employee Engagement',
  Analytics = 'Analytics & Reporting',
  Administration = 'Administration',
  Integration = 'Integration'
}


export interface IQuickStat {
  label: string;
  value: number;
  icon: string;
  color: string;
  link?: string;
}

export interface IRecentActivity {
  id: number;
  title: string;
  description: string;
  timestamp: Date;
  icon: string;
  iconColor: string;
  link?: string;
}
