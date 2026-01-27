// @ts-nocheck
/**
 * TaskMonitorService.ts
 *
 * Service for retrieving task monitoring data from SharePoint.
 * Calculates KPIs, escalations, department performance, and workload metrics
 * from JML_TaskAssignments and JML_Processes lists.
 *
 * @author JML Team
 * @version 1.0.0
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';
import { LoggingService } from './LoggingService';
import {
  ITaskMonitorKPI,
  IEscalation,
  IDepartmentPerformance,
  ITaskStreamItem,
  IEmployeeWorkload,
  ITaskHealthData
} from '../webparts/jmlTaskMonitor/components/IJmlTaskMonitorProps';

const logger = LoggingService.getInstance();

/**
 * Notification list name for in-app notifications
 */
const NOTIFICATIONS_LIST = 'JML_Notifications';

/**
 * Raw task assignment from SharePoint
 * Note: ProcessID is a Text field in JML_TaskAssignments, not a lookup
 */
interface ITaskAssignmentItem {
  Id: number;
  Title: string;
  Status: string;
  Priority: string;
  DueDate: string;
  CompletedDate?: string;
  AssignedTo?: { Title: string; EMail: string; Id: number };
  Department?: string;
  ProcessID?: string;  // Text field, not a lookup
  Category?: string;   // Process type indicator (Joiner/Mover/Leaver)
  BlockedReason?: string;
  EscalationLevel?: number;
  Created: string;
  Modified: string;
}

/**
 * Weekly trend data point
 */
export interface IWeeklyTrendData {
  day: string;
  completed: number;
  created: number;
}

/**
 * Task stream organized by grouping
 */
export interface ITaskStreamData {
  byStatus: {
    notStarted: ITaskStreamItem[];
    inProgress: ITaskStreamItem[];
    blocked: ITaskStreamItem[];
    overdue: ITaskStreamItem[];
  };
  byProcess: {
    joiner: ITaskStreamItem[];
    mover: ITaskStreamItem[];
    leaver: ITaskStreamItem[];
  };
  byDepartment: Record<string, ITaskStreamItem[]>;
  byTimeline: {
    overdue: ITaskStreamItem[];
    today: ITaskStreamItem[];
    tomorrow: ITaskStreamItem[];
    thisWeek: ITaskStreamItem[];
  };
}

/**
 * Complete task monitor data response
 */
export interface ITaskMonitorData {
  kpis: ITaskMonitorKPI[];
  escalations: IEscalation[];
  departments: IDepartmentPerformance[];
  taskStream: ITaskStreamData;
  workload: IEmployeeWorkload[];
  healthData: ITaskHealthData;
  weeklyTrend: IWeeklyTrendData[];
  lastUpdated: Date;
}

/**
 * Department configuration for icons and colors
 */
const DEPARTMENT_CONFIG: Record<string, { icon: string; iconBg: string }> = {
  'Human Resources': { icon: 'People', iconBg: '#deecf9' },
  'HR': { icon: 'People', iconBg: '#deecf9' },
  'IT': { icon: 'Devices4', iconBg: '#e0f7f8' },
  'Information Technology': { icon: 'Devices4', iconBg: '#e0f7f8' },
  'Finance': { icon: 'Money', iconBg: '#dff6dd' },
  'Legal': { icon: 'Certificate', iconBg: '#fff4ce' },
  'Security': { icon: 'Shield', iconBg: '#fde7e9' },
  'Facilities': { icon: 'CityNext', iconBg: '#e8daef' },
  'Operations': { icon: 'Settings', iconBg: '#fbeee6' }
};

/**
 * TaskMonitorService - Fetches and calculates task monitoring metrics
 */
export class TaskMonitorService {
  private readonly sp: SPFI;
  private readonly taskAssignmentsListName = 'JML_TaskAssignments';
  private readonly processesListName = 'JML_Processes';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all task monitor data in a single call
   */
  public async getTaskMonitorData(): Promise<ITaskMonitorData> {
    try {
      // Fetch all task assignments with expanded process data
      const tasks = await this.getTaskAssignments();

      // Calculate all metrics
      const healthData = this.calculateHealthData(tasks);
      const kpis = await this.calculateKPIs(tasks, healthData);
      const escalations = this.calculateEscalations(tasks);
      const departments = this.calculateDepartmentPerformance(tasks);
      const taskStream = this.organizeTaskStream(tasks);
      const workload = this.calculateWorkload(tasks);
      const weeklyTrend = await this.calculateWeeklyTrend();

      return {
        kpis,
        escalations,
        departments,
        taskStream,
        workload,
        healthData,
        weeklyTrend,
        lastUpdated: new Date()
      };
    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to get task monitor data', error);
      throw error;
    }
  }

  /**
   * Fetch task assignments from SharePoint
   */
  private async getTaskAssignments(): Promise<ITaskAssignmentItem[]> {
    try {
      // Query essential columns
      // NOTE: AssignedTo may be a text field (email) OR a Person lookup depending on list configuration
      // Try without expand first - if AssignedTo is a text field, expand will fail
      const items = await this.sp.web.lists
        .getByTitle(this.taskAssignmentsListName)
        .items
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'CompletedDate',
          'AssignedTo', 'AssignedToId',
          'Department', 'BlockedReason', 'Category', 'EscalationLevel',
          'Created', 'Modified'
        )
        .top(5000)();

      // Transform items - handle both text and lookup scenarios for AssignedTo
      return items.map((item: Record<string, unknown>) => {
        const assignedTo = item.AssignedTo;
        let assignedToObj: { Title: string; EMail: string; Id: number } | undefined;

        if (typeof assignedTo === 'string' && assignedTo) {
          // AssignedTo is a text field (email address)
          assignedToObj = {
            Title: assignedTo.split('@')[0] || assignedTo,
            EMail: assignedTo,
            Id: item.AssignedToId as number || 0
          };
        } else if (assignedTo && typeof assignedTo === 'object') {
          // AssignedTo is a Person lookup
          const personObj = assignedTo as { Title?: string; EMail?: string; Id?: number };
          assignedToObj = {
            Title: personObj.Title || 'Unknown',
            EMail: personObj.EMail || '',
            Id: personObj.Id || 0
          };
        }

        return {
          Id: item.Id as number,
          Title: item.Title as string,
          Status: item.Status as string,
          Priority: item.Priority as string,
          DueDate: item.DueDate as string,
          CompletedDate: item.CompletedDate as string | undefined,
          AssignedTo: assignedToObj,
          Department: item.Department as string | undefined,
          Category: item.Category as string | undefined,
          BlockedReason: item.BlockedReason as string | undefined,
          EscalationLevel: item.EscalationLevel as number | undefined,
          Created: item.Created as string,
          Modified: item.Modified as string
        } as ITaskAssignmentItem;
      });
    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to fetch task assignments', error);
      return [];
    }
  }

  /**
   * Calculate task health metrics
   */
  private calculateHealthData(tasks: ITaskAssignmentItem[]): ITaskHealthData {
    const now = new Date();

    const completed = tasks.filter(t => t.Status === 'Completed').length;
    const inProgress = tasks.filter(t => t.Status === 'In Progress').length;
    const notStarted = tasks.filter(t => t.Status === 'Not Started' || !t.Status).length;
    const overdue = tasks.filter(t => {
      if (t.Status === 'Completed') return false;
      if (!t.DueDate) return false;
      return new Date(t.DueDate) < now;
    }).length;

    return {
      total: tasks.length,
      completed,
      inProgress,
      notStarted,
      overdue
    };
  }

  /**
   * Calculate KPI metrics
   */
  private async calculateKPIs(
    tasks: ITaskAssignmentItem[],
    healthData: ITaskHealthData
  ): Promise<ITaskMonitorKPI[]> {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const endOfWeek = new Date(today);
    endOfWeek.setDate(endOfWeek.getDate() + 7);

    // Calculate metrics
    const activeTasks = tasks.filter(t => t.Status !== 'Completed').length;
    const overdueTasks = healthData.overdue;
    const dueThisWeek = tasks.filter(t => {
      if (t.Status === 'Completed') return false;
      if (!t.DueDate) return false;
      const dueDate = new Date(t.DueDate);
      return dueDate >= today && dueDate <= endOfWeek;
    }).length;
    const blockedTasks = tasks.filter(t => t.Status === 'Blocked').length;
    const completedToday = tasks.filter(t => {
      if (t.Status !== 'Completed' || !t.CompletedDate) return false;
      const completedDate = new Date(t.CompletedDate);
      return completedDate >= today;
    }).length;

    // Calculate SLA compliance (tasks completed on time / total completed)
    const completedTasks = tasks.filter(t => t.Status === 'Completed');
    const completedOnTime = completedTasks.filter(t => {
      if (!t.DueDate || !t.CompletedDate) return true;
      return new Date(t.CompletedDate) <= new Date(t.DueDate);
    }).length;
    const slaCompliance = completedTasks.length > 0
      ? Math.round((completedOnTime / completedTasks.length) * 1000) / 10
      : 100;

    // Calculate due today
    const dueToday = tasks.filter(t => {
      if (t.Status === 'Completed') return false;
      if (!t.DueDate) return false;
      const dueDate = new Date(t.DueDate);
      return dueDate.toDateString() === today.toDateString();
    }).length;

    return [
      {
        label: 'Total Active Tasks',
        value: activeTasks,
        icon: 'ClipboardList',
        trend: { direction: 'neutral', text: `${healthData.total} total` },
        variant: 'default'
      },
      {
        label: 'Overdue',
        value: overdueTasks,
        icon: 'ErrorBadge',
        trend: { direction: overdueTasks > 0 ? 'up' : 'neutral', text: 'Require attention' },
        variant: overdueTasks > 10 ? 'danger' : overdueTasks > 0 ? 'warning' : 'default'
      },
      {
        label: 'Due This Week',
        value: dueThisWeek,
        icon: 'Clock',
        trend: { direction: 'neutral', text: `${dueToday} due today` },
        variant: dueToday > 5 ? 'warning' : 'default'
      },
      {
        label: 'Blocked',
        value: blockedTasks,
        icon: 'Blocked',
        trend: { direction: blockedTasks > 0 ? 'up' : 'down', text: 'Awaiting resolution' },
        variant: blockedTasks > 5 ? 'warning' : 'default'
      },
      {
        label: 'Completed Today',
        value: completedToday,
        icon: 'CheckMark',
        trend: { direction: completedToday > 0 ? 'up' : 'neutral', text: 'Great progress!' },
        variant: 'success'
      },
      {
        label: 'SLA Compliance',
        value: slaCompliance,
        icon: 'Chart',
        trend: { direction: slaCompliance >= 90 ? 'up' : 'down', text: `${slaCompliance}% on-time` },
        variant: slaCompliance >= 90 ? 'success' : slaCompliance >= 80 ? 'warning' : 'danger'
      }
    ];
  }

  /**
   * Calculate escalations from overdue tasks
   */
  private calculateEscalations(tasks: ITaskAssignmentItem[]): IEscalation[] {
    const now = new Date();

    return tasks
      .filter(t => {
        if (t.Status === 'Completed') return false;
        if (!t.DueDate) return false;
        const dueDate = new Date(t.DueDate);
        const daysOverdue = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));
        return daysOverdue >= 3; // Only escalate tasks 3+ days overdue
      })
      .map(t => {
        const dueDate = new Date(t.DueDate);
        const daysOverdue = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));

        // Determine severity based on days overdue
        let severity: 'critical' | 'high' | 'medium';
        if (daysOverdue >= 7) severity = 'critical';
        else if (daysOverdue >= 5) severity = 'high';
        else severity = 'medium';

        // Determine escalation level
        const level = t.EscalationLevel || (daysOverdue >= 7 ? 2 : 1);

        return {
          id: t.Id.toString(),
          title: t.Title,
          employee: t.AssignedTo?.Title || 'Unknown',
          severity,
          daysOverdue,
          department: t.Department || 'Unknown',
          processType: t.Category || 'Task',  // Use Category field instead of Process lookup
          escalatedAt: new Date(t.Modified),
          level
        };
      })
      .sort((a, b) => b.daysOverdue - a.daysOverdue)
      .slice(0, 10); // Top 10 escalations
  }

  /**
   * Calculate department performance metrics
   */
  private calculateDepartmentPerformance(tasks: ITaskAssignmentItem[]): IDepartmentPerformance[] {
    // Group tasks by department
    const deptMap = new Map<string, ITaskAssignmentItem[]>();

    tasks.forEach(t => {
      const dept = t.Department || 'Other';
      if (!deptMap.has(dept)) {
        deptMap.set(dept, []);
      }
      deptMap.get(dept)!.push(t);
    });

    const now = new Date();

    return Array.from(deptMap.entries())
      .map(([name, deptTasks]) => {
        const activeCount = deptTasks.filter(t => t.Status !== 'Completed').length;
        const completedTasks = deptTasks.filter(t => t.Status === 'Completed');
        const completionRate = deptTasks.length > 0
          ? Math.round((completedTasks.length / deptTasks.length) * 100)
          : 0;

        // Calculate average completion time
        const completedWithDates = completedTasks.filter(t => t.Created && t.CompletedDate);
        let avgTime = 'N/A';
        if (completedWithDates.length > 0) {
          const totalDays = completedWithDates.reduce((sum, t) => {
            const created = new Date(t.Created);
            const completed = new Date(t.CompletedDate!);
            return sum + (completed.getTime() - created.getTime()) / (1000 * 60 * 60 * 24);
          }, 0);
          avgTime = `${(totalDays / completedWithDates.length).toFixed(1)} days`;
        }

        const overdueCount = deptTasks.filter(t => {
          if (t.Status === 'Completed') return false;
          if (!t.DueDate) return false;
          return new Date(t.DueDate) < now;
        }).length;

        // Determine status based on metrics
        let status: 'good' | 'warning' | 'danger';
        if (completionRate >= 85 && overdueCount <= 3) status = 'good';
        else if (completionRate >= 70 || overdueCount <= 10) status = 'warning';
        else status = 'danger';

        const config = DEPARTMENT_CONFIG[name] || { icon: 'Org', iconBg: '#f3f2f1' };

        return {
          name,
          icon: config.icon,
          iconBg: config.iconBg,
          activeCount,
          completionRate,
          avgTime,
          overdueCount,
          status
        };
      })
      .filter(d => d.activeCount > 0 || d.overdueCount > 0)
      .sort((a, b) => b.activeCount - a.activeCount)
      .slice(0, 6); // Top 6 departments
  }

  /**
   * Organize tasks into stream views
   */
  private organizeTaskStream(tasks: ITaskAssignmentItem[]): ITaskStreamData {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const endOfWeek = new Date(today);
    endOfWeek.setDate(endOfWeek.getDate() + 7);

    const mapToStreamItem = (t: ITaskAssignmentItem): ITaskStreamItem => {
      const dueDate = t.DueDate ? new Date(t.DueDate) : null;
      const isOverdue = dueDate ? dueDate < now && t.Status !== 'Completed' : false;
      const isDueSoon = dueDate ? dueDate <= tomorrow && !isOverdue : false;

      let dueText = 'No due date';
      if (dueDate) {
        if (isOverdue) {
          const daysOverdue = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));
          dueText = `${daysOverdue} days overdue`;
        } else if (dueDate.toDateString() === today.toDateString()) {
          dueText = 'Due today';
        } else if (dueDate.toDateString() === tomorrow.toDateString()) {
          dueText = 'Due tomorrow';
        } else {
          const daysUntil = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
          dueText = `Due in ${daysUntil} days`;
        }
      }

      const assigneeName = t.AssignedTo?.Title || 'Unassigned';
      const initials = assigneeName.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2);

      return {
        id: t.Id.toString(),
        title: t.Title,
        processId: `TASK-${t.Id}`,
        processType: t.Category || 'Task',
        priority: (t.Priority?.toLowerCase() as 'high' | 'medium' | 'normal') || 'normal',
        assignee: assigneeName,
        assigneeInitials: initials,
        dueText,
        isOverdue,
        isDueSoon,
        blockedReason: t.BlockedReason
      };
    };

    const activeTasks = tasks.filter(t => t.Status !== 'Completed');
    const streamItems = activeTasks.map(mapToStreamItem);

    // By Status
    const byStatus = {
      notStarted: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        return task?.Status === 'Not Started' || !task?.Status;
      }).slice(0, 10),
      inProgress: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        return task?.Status === 'In Progress';
      }).slice(0, 10),
      blocked: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        return task?.Status === 'Blocked';
      }).slice(0, 10),
      overdue: streamItems.filter(t => t.isOverdue).slice(0, 10)
    };

    // By Process Type
    const byProcess = {
      joiner: streamItems.filter(t => t.processType === 'Joiner').slice(0, 10),
      mover: streamItems.filter(t => t.processType === 'Mover').slice(0, 10),
      leaver: streamItems.filter(t => t.processType === 'Leaver').slice(0, 10)
    };

    // By Department
    const byDepartment: Record<string, ITaskStreamItem[]> = {};
    activeTasks.forEach(t => {
      const dept = t.Department || 'Other';
      if (!byDepartment[dept]) {
        byDepartment[dept] = [];
      }
      const item = streamItems.find(s => s.id === t.Id.toString());
      if (item && byDepartment[dept].length < 10) {
        byDepartment[dept].push(item);
      }
    });

    // By Timeline
    const byTimeline = {
      overdue: streamItems.filter(t => t.isOverdue).slice(0, 10),
      today: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        if (!task?.DueDate) return false;
        const dueDate = new Date(task.DueDate);
        return dueDate.toDateString() === today.toDateString() && !t.isOverdue;
      }).slice(0, 10),
      tomorrow: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        if (!task?.DueDate) return false;
        const dueDate = new Date(task.DueDate);
        return dueDate.toDateString() === tomorrow.toDateString();
      }).slice(0, 10),
      thisWeek: streamItems.filter(t => {
        const task = tasks.find(x => x.Id.toString() === t.id);
        if (!task?.DueDate) return false;
        const dueDate = new Date(task.DueDate);
        return dueDate > tomorrow && dueDate <= endOfWeek;
      }).slice(0, 10)
    };

    return { byStatus, byProcess, byDepartment, byTimeline };
  }

  /**
   * Calculate employee workload metrics
   */
  private calculateWorkload(tasks: ITaskAssignmentItem[]): IEmployeeWorkload[] {
    const now = new Date();
    const workloadMap = new Map<string, {
      name: string;
      department: string;
      active: number;
      overdue: number;
    }>();

    tasks
      .filter(t => t.Status !== 'Completed' && t.AssignedTo?.Title)
      .forEach(t => {
        const userId = t.AssignedTo!.Id.toString();
        const name = t.AssignedTo!.Title;

        if (!workloadMap.has(userId)) {
          workloadMap.set(userId, {
            name,
            department: t.Department || 'Unknown',
            active: 0,
            overdue: 0
          });
        }

        const entry = workloadMap.get(userId)!;
        entry.active++;

        if (t.DueDate && new Date(t.DueDate) < now) {
          entry.overdue++;
        }
      });

    // Calculate max active tasks for percentage
    const entries = Array.from(workloadMap.entries());
    const maxActive = Math.max(...entries.map(([, e]) => e.active), 1);

    return entries
      .map(([id, data]) => {
        const workloadPercent = Math.round((data.active / maxActive) * 100);
        const initials = data.name.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2);

        return {
          id,
          name: data.name,
          initials,
          department: data.department,
          activeCount: data.active,
          overdueCount: data.overdue,
          workloadPercent,
          isOverloaded: workloadPercent >= 80 || data.overdue >= 5
        };
      })
      .sort((a, b) => b.workloadPercent - a.workloadPercent)
      .slice(0, 8); // Top 8 employees by workload
  }

  /**
   * Calculate weekly trend data
   */
  private async calculateWeeklyTrend(): Promise<IWeeklyTrendData[]> {
    const now = new Date();
    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const trend: IWeeklyTrendData[] = [];

    try {
      // Get tasks from the last 7 days
      const startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 7);

      const recentTasks = await this.sp.web.lists
        .getByTitle(this.taskAssignmentsListName)
        .items
        .select('Id', 'Status', 'CompletedDate', 'Created')
        .filter(`Created ge datetime'${startDate.toISOString()}' or CompletedDate ge datetime'${startDate.toISOString()}'`)
        .top(5000)();

      // Group by day
      for (let i = 6; i >= 0; i--) {
        const date = new Date(now);
        date.setDate(date.getDate() - i);
        const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        const dayEnd = new Date(dayStart);
        dayEnd.setDate(dayEnd.getDate() + 1);

        const completed = recentTasks.filter(t => {
          if (!t.CompletedDate) return false;
          const completedDate = new Date(t.CompletedDate);
          return completedDate >= dayStart && completedDate < dayEnd;
        }).length;

        const created = recentTasks.filter(t => {
          const createdDate = new Date(t.Created);
          return createdDate >= dayStart && createdDate < dayEnd;
        }).length;

        trend.push({
          day: days[date.getDay()],
          completed,
          created
        });
      }
    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to calculate weekly trend', error);
      // Return empty trend data
      for (let i = 6; i >= 0; i--) {
        const date = new Date(now);
        date.setDate(date.getDate() - i);
        trend.push({
          day: days[date.getDay()],
          completed: 0,
          created: 0
        });
      }
    }

    return trend;
  }

  /**
   * Send reminder notifications for escalated tasks
   */
  public async sendReminders(taskIds: string[], message: string): Promise<boolean> {
    try {
      // In a real implementation, this would:
      // 1. Get task details
      // 2. Create notification records in JML_Notifications
      // 3. Trigger Power Automate flow for email
      logger.error('TaskMonitorService', `Sending reminders for ${taskIds.length} tasks`);

      // For now, just log the action
      console.log('Reminders sent:', { taskIds, message });
      return true;
    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to send reminders', error);
      return false;
    }
  }

  /**
   * Reassign tasks to a different user
   */
  public async reassignTasks(taskIds: string[], newAssigneeId: number): Promise<boolean> {
    try {
      for (const taskId of taskIds) {
        await this.sp.web.lists
          .getByTitle(this.taskAssignmentsListName)
          .items.getById(parseInt(taskId, 10))
          .update({
            AssignedToId: newAssigneeId
          });
      }
      return true;
    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to reassign tasks', error);
      return false;
    }
  }

  /**
   * Run escalation check on all tasks
   * Evaluates tasks against escalation rules and returns count of escalated tasks
   * @param sendNotifications - If true, sends in-app notifications for escalated tasks
   */
  public async runEscalationCheck(sendNotifications: boolean = true): Promise<IEscalationCheckResult> {
    try {
      logger.info('TaskMonitorService', 'Running manual escalation check');

      // Get all task assignments
      const tasks = await this.getTaskAssignments();

      // Get escalation rules from JML_TaskEscalationRules
      const rules = await this.getEscalationRules();

      if (rules.length === 0) {
        return {
          success: true,
          tasksEvaluated: tasks.length,
          tasksEscalated: 0,
          message: 'No active escalation rules found. Please create escalation rules first.'
        };
      }

      const now = new Date();
      const escalatedTasks: IEscalatedTaskInfo[] = [];

      // Evaluate each task against rules
      for (const task of tasks) {
        // Skip completed or cancelled tasks
        if (task.Status === 'Completed' || task.Status === 'Cancelled' || task.Status === 'Skipped') {
          continue;
        }

        const taskDueDate = task.DueDate ? new Date(task.DueDate) : null;
        const taskCreated = task.Created ? new Date(task.Created) : null;

        for (const rule of rules) {
          // Check if rule applies to this task
          if (!this.ruleAppliesToTask(rule, task)) {
            continue;
          }

          let shouldEscalate = false;
          let triggerReason = '';

          switch (rule.EscalationTrigger) {
            case 'OverdueBy':
              if (taskDueDate && taskDueDate < now) {
                const hoursOverdue = (now.getTime() - taskDueDate.getTime()) / (1000 * 60 * 60);
                if (hoursOverdue >= rule.TriggerValue) {
                  shouldEscalate = true;
                  triggerReason = `Overdue by ${Math.round(hoursOverdue)} hours`;
                }
              }
              break;

            case 'NotStartedAfter':
              if (task.Status === 'Not Started' && taskCreated) {
                const hoursNotStarted = (now.getTime() - taskCreated.getTime()) / (1000 * 60 * 60);
                if (hoursNotStarted >= rule.TriggerValue) {
                  shouldEscalate = true;
                  triggerReason = `Not started after ${Math.round(hoursNotStarted)} hours`;
                }
              }
              break;

            case 'ApproachingDue':
              if (taskDueDate && taskDueDate > now) {
                const hoursUntilDue = (taskDueDate.getTime() - now.getTime()) / (1000 * 60 * 60);
                if (hoursUntilDue <= rule.TriggerValue) {
                  shouldEscalate = true;
                  triggerReason = `Due in ${Math.round(hoursUntilDue)} hours`;
                }
              }
              break;

            case 'HighPriorityOverdue':
              if ((task.Priority === 'High' || task.Priority === 'Critical') && taskDueDate && taskDueDate < now) {
                const hoursOverdue = (now.getTime() - taskDueDate.getTime()) / (1000 * 60 * 60);
                if (hoursOverdue >= rule.TriggerValue) {
                  shouldEscalate = true;
                  triggerReason = `High priority overdue by ${Math.round(hoursOverdue)} hours`;
                }
              }
              break;

            case 'StuckInStatus':
              // Check Modified date to see if task has been stuck
              const taskModified = task.Modified ? new Date(task.Modified) : null;
              if (taskModified) {
                const hoursStuck = (now.getTime() - taskModified.getTime()) / (1000 * 60 * 60);
                if (hoursStuck >= rule.TriggerValue) {
                  shouldEscalate = true;
                  triggerReason = `Stuck in ${task.Status} for ${Math.round(hoursStuck)} hours`;
                }
              }
              break;
          }

          if (shouldEscalate) {
            // Check if already escalated at this level
            const currentLevel = task.EscalationLevel || 0;
            if (rule.EscalationLevel > currentLevel) {
              escalatedTasks.push({
                taskId: task.Id,
                taskTitle: task.Title,
                assignee: task.AssignedTo?.Title || 'Unassigned',
                assigneeId: task.AssignedTo?.Id,
                ruleTitle: rule.Title,
                escalationLevel: rule.EscalationLevel,
                triggerReason,
                notifyRoles: rule.NotifyRoles
              });

              // Update escalation level on the task
              await this.sp.web.lists
                .getByTitle(this.taskAssignmentsListName)
                .items.getById(task.Id)
                .update({
                  EscalationLevel: rule.EscalationLevel,
                  EscalationSent: true
                });

              // Send notification if enabled
              if (sendNotifications && task.AssignedTo?.Id) {
                await this.sendEscalationNotification(task, rule, triggerReason);
              }
            }
          }
        }
      }

      logger.info('TaskMonitorService', `Escalation check complete: ${escalatedTasks.length} tasks escalated`);

      return {
        success: true,
        tasksEvaluated: tasks.length,
        tasksEscalated: escalatedTasks.length,
        escalatedTasks,
        message: escalatedTasks.length > 0
          ? `${escalatedTasks.length} tasks have been escalated based on ${rules.length} active rules.`
          : `All ${tasks.length} tasks evaluated. No escalations required.`
      };

    } catch (error) {
      logger.error('TaskMonitorService', 'Failed to run escalation check', error);
      return {
        success: false,
        tasksEvaluated: 0,
        tasksEscalated: 0,
        message: `Escalation check failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  /**
   * Get active escalation rules from SharePoint
   */
  private async getEscalationRules(): Promise<IEscalationRule[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('JML_TaskEscalationRules')
        .items
        .filter("IsActive eq 1")
        .select(
          'Id', 'Title', 'IsGlobalRule', 'EscalationTrigger', 'TriggerValue',
          'EscalationLevel', 'NotifyRoles', 'ApplicesToDepartment', 'AppliesToCategoryFilter',
          'AutoReassign', 'ReassignToRole'
        )
        .orderBy('EscalationLevel')();

      return items.map(item => ({
        Id: item.Id,
        Title: item.Title,
        IsGlobalRule: item.IsGlobalRule === true,
        EscalationTrigger: item.EscalationTrigger,
        TriggerValue: item.TriggerValue || 0,
        EscalationLevel: item.EscalationLevel || 1,
        NotifyRoles: item.NotifyRoles ? JSON.parse(item.NotifyRoles) : [],
        AppliesToDepartment: item.ApplicesToDepartment,
        AppliesToCategoryFilter: item.AppliesToCategoryFilter ? JSON.parse(item.AppliesToCategoryFilter) : [],
        AutoReassign: item.AutoReassign === true,
        ReassignToRole: item.ReassignToRole
      }));
    } catch (error) {
      logger.warn('TaskMonitorService', 'Failed to get escalation rules - list may not exist', error);
      return [];
    }
  }

  /**
   * Check if an escalation rule applies to a specific task
   */
  private ruleAppliesToTask(rule: IEscalationRule, task: ITaskAssignmentItem): boolean {
    // Global rules apply to all tasks
    if (rule.IsGlobalRule) {
      return true;
    }

    // Check department filter
    if (rule.AppliesToDepartment && task.Department !== rule.AppliesToDepartment) {
      return false;
    }

    // Check category filter
    if (rule.AppliesToCategoryFilter && rule.AppliesToCategoryFilter.length > 0) {
      if (!task.Category || !rule.AppliesToCategoryFilter.includes(task.Category)) {
        return false;
      }
    }

    return true;
  }

  /**
   * Send escalation notification to JML_Notifications list
   */
  private async sendEscalationNotification(
    task: ITaskAssignmentItem,
    rule: IEscalationRule,
    triggerReason: string
  ): Promise<void> {
    try {
      // Determine notification recipients
      const recipientIds: number[] = [];

      // Always notify assignee
      if (task.AssignedTo?.Id) {
        recipientIds.push(task.AssignedTo.Id);
      }

      // Get additional recipients based on roles
      // For now, we notify the assignee; in production you'd look up role assignments
      // (Manager, DepartmentHead, etc.) from configuration lists

      // Create notification message based on escalation level
      const levelText = rule.EscalationLevel === 1 ? 'Warning' :
                        rule.EscalationLevel === 2 ? 'Urgent' :
                        rule.EscalationLevel === 3 ? 'Critical' : 'Alert';

      const notificationTitle = `[${levelText}] Task Escalation: ${task.Title}`;
      const notificationBody = `Your task "${task.Title}" has been escalated.\n\n` +
        `Reason: ${triggerReason}\n` +
        `Rule: ${rule.Title}\n` +
        `Escalation Level: ${rule.EscalationLevel}\n\n` +
        `Please address this task as soon as possible.`;

      // Create notification record for each recipient
      for (const recipientId of recipientIds) {
        try {
          await this.sp.web.lists.getByTitle(NOTIFICATIONS_LIST).items.add({
            Title: notificationTitle,
            NotificationType: 'Escalation',
            Message: notificationBody,
            RecipientId: recipientId,
            RelatedItemType: 'TaskAssignment',
            RelatedItemId: task.Id,
            Priority: rule.EscalationLevel >= 2 ? 'High' : 'Medium',
            IsRead: false,
            RequiresAction: true,
            ActionUrl: `/sites/JML/SitePages/MyTasks.aspx?taskId=${task.Id}`
          });
        } catch (notifyError) {
          logger.warn('TaskMonitorService', `Failed to create notification for user ${recipientId}`, notifyError);
        }
      }

      logger.info('TaskMonitorService', `Sent escalation notifications for task ${task.Id} to ${recipientIds.length} recipients`);
    } catch (error) {
      logger.warn('TaskMonitorService', `Failed to send escalation notification for task ${task.Id}`, error);
      // Don't throw - notification failure shouldn't fail the escalation check
    }
  }
}

/**
 * Escalation rule from JML_TaskEscalationRules
 */
interface IEscalationRule {
  Id: number;
  Title: string;
  IsGlobalRule: boolean;
  EscalationTrigger: string;
  TriggerValue: number;
  EscalationLevel: number;
  NotifyRoles: string[];
  AppliesToDepartment?: string;
  AppliesToCategoryFilter?: string[];
  AutoReassign: boolean;
  ReassignToRole?: string;
}

/**
 * Result of escalation check
 */
export interface IEscalationCheckResult {
  success: boolean;
  tasksEvaluated: number;
  tasksEscalated: number;
  escalatedTasks?: IEscalatedTaskInfo[];
  message: string;
}

/**
 * Info about an escalated task
 */
export interface IEscalatedTaskInfo {
  taskId: number;
  taskTitle: string;
  assignee: string;
  assigneeId?: number;
  ruleTitle: string;
  escalationLevel: number;
  triggerReason: string;
  notifyRoles?: string[];
}

export default TaskMonitorService;
