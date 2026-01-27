// @ts-nocheck
/* eslint-disable */
/**
 * Default Workflow Definitions for JML Solution
 *
 * These workflow definitions are the standard templates for Joiner, Mover, and Leaver processes.
 * They should be seeded into the JML_WorkflowDefinitions SharePoint list during deployment.
 *
 * Usage:
 * 1. Run the provisioning script: scripts/Provision-WorkflowDefinitions.ps1
 * 2. Or manually create items in JML_WorkflowDefinitions list using these JSON structures
 */

import {
  IWorkflowStep,
  IWorkflowVariable,
  StepType,
  TransitionType,
  ActionType,
  ConditionOperator
} from '../../models/IWorkflow';
import { ProcessType } from '../../models/ICommon';

// ============================================================================
// JOINER (ONBOARDING) WORKFLOW DEFINITION
// ============================================================================

export const JOINER_WORKFLOW_STEPS: IWorkflowStep[] = [
  // Step 1: Start
  {
    id: 'STEP-START',
    name: 'Start Onboarding',
    description: 'Initialize the onboarding workflow',
    type: StepType.Start,
    order: 1,
    config: {},
    onComplete: { type: TransitionType.Next }
  },

  // Step 2: Notify Manager
  {
    id: 'STEP-NOTIFY-MANAGER',
    name: 'Notify Manager',
    description: 'Send notification to the manager about the new hire',
    type: StepType.Notification,
    order: 2,
    config: {
      notificationType: 'ProcessStarted',
      recipientField: 'managerId',
      notificationSubject: 'New Hire Onboarding Started: {{employeeName}}',
      messageTemplate: 'A new onboarding process has been started for {{employeeName}} in {{department}}. Start date: {{startDate}}. Please review and complete your assigned tasks.'
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 4,
      breachHours: 8
    }
  },

  // Step 3: Pre-Arrival Tasks
  {
    id: 'STEP-PREARRRIVAL-TASKS',
    name: 'Pre-Arrival Tasks',
    description: 'Assign pre-arrival tasks (IT setup, workspace, accounts)',
    type: StepType.AssignTasks,
    order: 3,
    config: {
      taskTitle: 'Pre-Arrival Setup',
      assigneeRole: 'HR Admin',
      dueDaysFromNow: -3 // 3 days before start date
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 48,
      breachHours: 72,
      escalateTo: 'HR Manager'
    }
  },

  // Step 4: Wait for Pre-Arrival Tasks
  {
    id: 'STEP-WAIT-PREARRRIVAL',
    name: 'Wait for Pre-Arrival Completion',
    description: 'Wait for all pre-arrival tasks to be completed',
    type: StepType.WaitForTasks,
    order: 4,
    config: {
      waitForTaskIds: ['STEP-PREARRRIVAL-TASKS'],
      waitCondition: 'all'
    },
    onComplete: { type: TransitionType.Next },
    timeoutHours: 168, // 7 days
    onTimeout: {
      type: TransitionType.Goto,
      targetStepId: 'STEP-ESCALATE-PREARRRIVAL'
    }
  },

  // Step 4b: Escalation for Pre-Arrival (conditional)
  {
    id: 'STEP-ESCALATE-PREARRRIVAL',
    name: 'Escalate Pre-Arrival Tasks',
    description: 'Escalate overdue pre-arrival tasks',
    type: StepType.Notification,
    order: 5,
    config: {
      notificationType: 'TaskEscalated',
      recipientRole: 'HR Manager',
      notificationSubject: 'ESCALATION: Pre-Arrival Tasks Overdue for {{employeeName}}',
      messageTemplate: 'Pre-arrival tasks for {{employeeName}} are overdue. Immediate attention required. Start date: {{startDate}}.'
    },
    conditions: [
      {
        id: 'cond-timeout',
        field: 'step.timedOut',
        operator: ConditionOperator.Equals,
        value: true
      }
    ],
    onComplete: {
      type: TransitionType.Goto,
      targetStepId: 'STEP-WAIT-PREARRRIVAL'
    }
  },

  // Step 5: Day 1 Welcome
  {
    id: 'STEP-DAY1-WELCOME',
    name: 'Day 1 Welcome',
    description: 'Send welcome notification and assign Day 1 tasks',
    type: StepType.Notification,
    order: 6,
    config: {
      notificationType: 'Welcome',
      recipientField: 'employeeEmail',
      notificationSubject: 'Welcome to the Team, {{employeeName}}!',
      messageTemplate: 'Welcome to {{department}}! Your onboarding journey begins today. Please check your tasks in the My Tasks page for your first day activities.'
    },
    onComplete: { type: TransitionType.Next },
    conditions: [
      {
        id: 'cond-start-date',
        field: 'process.startDate',
        operator: ConditionOperator.LessThanOrEqual,
        value: '{{today}}'
      }
    ]
  },

  // Step 6: Day 1 Tasks
  {
    id: 'STEP-DAY1-TASKS',
    name: 'Day 1 Tasks',
    description: 'Assign Day 1 onboarding tasks',
    type: StepType.AssignTasks,
    order: 7,
    config: {
      taskTitle: 'Day 1 Onboarding Tasks',
      assigneeField: 'employeeId',
      dueDaysFromNow: 1
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 8,
      breachHours: 24
    }
  },

  // Step 7: Manager Introduction Meeting
  {
    id: 'STEP-MANAGER-MEETING',
    name: 'Manager Introduction',
    description: 'Schedule manager introduction meeting',
    type: StepType.CreateTask,
    order: 8,
    config: {
      taskTitle: 'Conduct introduction meeting with {{employeeName}}',
      assigneeField: 'managerId',
      dueDaysFromNow: 1
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 8: Week 1 Tasks
  {
    id: 'STEP-WEEK1-TASKS',
    name: 'Week 1 Tasks',
    description: 'Assign Week 1 training and orientation tasks',
    type: StepType.AssignTasks,
    order: 9,
    config: {
      taskTitle: 'Week 1 Training & Orientation',
      assigneeField: 'employeeId',
      dueDaysFromNow: 7
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 120,
      breachHours: 168
    }
  },

  // Step 9: Wait for Week 1 Completion
  {
    id: 'STEP-WAIT-WEEK1',
    name: 'Wait for Week 1 Completion',
    description: 'Wait for Week 1 tasks to complete',
    type: StepType.WaitForTasks,
    order: 10,
    config: {
      waitForTaskIds: ['STEP-WEEK1-TASKS'],
      waitCondition: 'all'
    },
    onComplete: { type: TransitionType.Next },
    timeoutHours: 240 // 10 days
  },

  // Step 10: Week 1 Check-in
  {
    id: 'STEP-WEEK1-CHECKIN',
    name: 'Week 1 Check-in',
    description: 'Manager check-in after first week',
    type: StepType.CreateTask,
    order: 11,
    config: {
      taskTitle: 'Week 1 check-in with {{employeeName}}',
      assigneeField: 'managerId',
      dueDaysFromNow: 7
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 11: Month 1 Tasks
  {
    id: 'STEP-MONTH1-TASKS',
    name: 'Month 1 Tasks',
    description: 'Assign Month 1 tasks and training',
    type: StepType.AssignTasks,
    order: 12,
    config: {
      taskTitle: 'Month 1 Tasks',
      assigneeField: 'employeeId',
      dueDaysFromNow: 30
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 600,
      breachHours: 720
    }
  },

  // Step 12: Wait for Month 1 Completion
  {
    id: 'STEP-WAIT-MONTH1',
    name: 'Wait for Month 1 Completion',
    description: 'Wait for Month 1 tasks to complete',
    type: StepType.WaitForTasks,
    order: 13,
    config: {
      waitForTaskIds: ['STEP-MONTH1-TASKS'],
      waitCondition: 'all'
    },
    onComplete: { type: TransitionType.Next },
    timeoutHours: 960 // 40 days
  },

  // Step 13: 30-Day Review
  {
    id: 'STEP-30DAY-REVIEW',
    name: '30-Day Review',
    description: 'Manager conducts 30-day performance review',
    type: StepType.CreateTask,
    order: 14,
    config: {
      taskTitle: 'Conduct 30-day review for {{employeeName}}',
      assigneeField: 'managerId',
      dueDaysFromNow: 35
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 14: Probation Check (Condition)
  {
    id: 'STEP-PROBATION-CHECK',
    name: 'Probation Period Check',
    description: 'Check if extended probation tasks are needed',
    type: StepType.Condition,
    order: 15,
    config: {
      conditionGroups: [
        {
          logic: 'AND',
          conditions: [
            {
              id: 'cond-probation',
              field: 'process.requiresProbationExtension',
              operator: ConditionOperator.Equals,
              value: true
            }
          ]
        }
      ]
    },
    onComplete: {
      type: TransitionType.Branch,
      branches: [
        {
          name: 'Extended Probation',
          conditions: [
            {
              logic: 'AND',
              conditions: [
                {
                  id: 'branch-extended',
                  field: 'process.requiresProbationExtension',
                  operator: ConditionOperator.Equals,
                  value: true
                }
              ]
            }
          ],
          targetStepId: 'STEP-EXTENDED-PROBATION'
        },
        {
          name: 'Standard Completion',
          conditions: [],
          targetStepId: 'STEP-90DAY-REVIEW',
          isDefault: true
        }
      ]
    }
  },

  // Step 15: Extended Probation Tasks (conditional branch)
  {
    id: 'STEP-EXTENDED-PROBATION',
    name: 'Extended Probation Tasks',
    description: 'Additional tasks for extended probation',
    type: StepType.AssignTasks,
    order: 16,
    config: {
      taskTitle: 'Extended Probation Tasks',
      assigneeField: 'employeeId',
      dueDaysFromNow: 60
    },
    onComplete: {
      type: TransitionType.Goto,
      targetStepId: 'STEP-90DAY-REVIEW'
    }
  },

  // Step 16: 90-Day Review
  {
    id: 'STEP-90DAY-REVIEW',
    name: '90-Day Review',
    description: 'Manager conducts 90-day performance review',
    type: StepType.CreateTask,
    order: 17,
    config: {
      taskTitle: 'Conduct 90-day review for {{employeeName}}',
      assigneeField: 'managerId',
      dueDaysFromNow: 95
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 17: HR Approval for Probation Completion
  {
    id: 'STEP-HR-APPROVAL',
    name: 'HR Approval',
    description: 'HR approves successful completion of probation',
    type: StepType.Approval,
    order: 18,
    config: {
      approverRole: 'HR Admin',
      approvalTemplateId: undefined // Will use default
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 48,
      breachHours: 72
    }
  },

  // Step 18: Completion Notification
  {
    id: 'STEP-COMPLETION-NOTIFY',
    name: 'Onboarding Complete Notification',
    description: 'Notify all stakeholders of successful onboarding completion',
    type: StepType.Notification,
    order: 19,
    config: {
      notificationType: 'WorkflowCompleted',
      recipientField: 'employeeEmail',
      notificationSubject: 'Congratulations! Onboarding Complete - {{employeeName}}',
      messageTemplate: 'Congratulations {{employeeName}}! Your onboarding has been successfully completed. Welcome to the team!'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 19: End
  {
    id: 'STEP-END',
    name: 'End Onboarding',
    description: 'Complete the onboarding workflow',
    type: StepType.End,
    order: 20,
    config: {},
    onComplete: { type: TransitionType.End }
  }
];

export const JOINER_WORKFLOW_VARIABLES: IWorkflowVariable[] = [
  { name: 'employeeName', type: 'string', description: 'Name of the new employee' },
  { name: 'employeeEmail', type: 'string', description: 'Email of the new employee' },
  { name: 'department', type: 'string', description: 'Department the employee is joining' },
  { name: 'managerId', type: 'number', description: 'Manager SharePoint user ID' },
  { name: 'startDate', type: 'date', description: 'Employee start date' },
  { name: 'requiresProbationExtension', type: 'boolean', defaultValue: false, description: 'Flag for extended probation' },
  { name: 'buddyId', type: 'number', description: 'Assigned buddy user ID' }
];

export const DEFAULT_JOINER_WORKFLOW = {
  Title: 'Standard Onboarding Workflow',
  WorkflowCode: 'WF-JOINER-STD-001',
  Description: 'Standard onboarding workflow for new employees. Covers pre-arrival through 90-day probation completion.',
  Version: '1.0.0',
  ProcessType: ProcessType.Joiner,
  IsActive: true,
  IsDefault: true,
  Category: 'Onboarding',
  Tags: 'onboarding,joiner,standard,probation',
  EstimatedDuration: 2160, // 90 days in hours
  Steps: JSON.stringify(JOINER_WORKFLOW_STEPS),
  Variables: JSON.stringify(JOINER_WORKFLOW_VARIABLES),
  TriggerConditions: JSON.stringify([])
};

// ============================================================================
// MOVER (TRANSFER/RELOCATION) WORKFLOW DEFINITION
// ============================================================================

export const MOVER_WORKFLOW_STEPS: IWorkflowStep[] = [
  // Step 1: Start
  {
    id: 'STEP-START',
    name: 'Start Transfer',
    description: 'Initialize the transfer/relocation workflow',
    type: StepType.Start,
    order: 1,
    config: {},
    onComplete: { type: TransitionType.Next }
  },

  // Step 2: Notify Current Manager
  {
    id: 'STEP-NOTIFY-CURRENT-MGR',
    name: 'Notify Current Manager',
    description: 'Inform current manager about the transfer',
    type: StepType.Notification,
    order: 2,
    config: {
      notificationType: 'ProcessStarted',
      recipientField: 'currentManagerId',
      notificationSubject: 'Employee Transfer: {{employeeName}} - Action Required',
      messageTemplate: '{{employeeName}} is transferring from your team. Please complete knowledge transfer tasks and update access permissions. Effective date: {{transferDate}}.'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 3: Notify New Manager
  {
    id: 'STEP-NOTIFY-NEW-MGR',
    name: 'Notify New Manager',
    description: 'Inform new manager about the incoming transfer',
    type: StepType.Notification,
    order: 3,
    config: {
      notificationType: 'ProcessStarted',
      recipientField: 'newManagerId',
      notificationSubject: 'New Team Member: {{employeeName}} Transferring',
      messageTemplate: '{{employeeName}} will be joining your team from {{currentDepartment}}. Transfer effective: {{transferDate}}. Please prepare onboarding tasks for their new role.'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 4: Knowledge Transfer Tasks
  {
    id: 'STEP-KNOWLEDGE-TRANSFER',
    name: 'Knowledge Transfer',
    description: 'Current team knowledge transfer tasks',
    type: StepType.AssignTasks,
    order: 4,
    config: {
      taskTitle: 'Complete knowledge transfer documentation',
      assigneeField: 'employeeId',
      dueDaysFromNow: 7
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 120,
      breachHours: 168
    }
  },

  // Step 5: Current Manager Handover Approval
  {
    id: 'STEP-HANDOVER-APPROVAL',
    name: 'Handover Approval',
    description: 'Current manager approves knowledge transfer completion',
    type: StepType.Approval,
    order: 5,
    config: {
      approverField: 'currentManagerId'
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 24,
      breachHours: 48
    }
  },

  // Step 6: IT Access Update
  {
    id: 'STEP-IT-ACCESS-UPDATE',
    name: 'IT Access Update',
    description: 'Update system access for new role/location',
    type: StepType.AssignTasks,
    order: 6,
    config: {
      taskTitle: 'Update IT access for {{employeeName}} role change',
      assigneeRole: 'IT Admin',
      dueDaysFromNow: 3
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 48,
      breachHours: 72
    }
  },

  // Step 7: Check if Location Change
  {
    id: 'STEP-LOCATION-CHECK',
    name: 'Location Change Check',
    description: 'Check if physical relocation is needed',
    type: StepType.Condition,
    order: 7,
    config: {
      conditionGroups: [
        {
          logic: 'AND',
          conditions: [
            {
              id: 'cond-location',
              field: 'process.isLocationChange',
              operator: ConditionOperator.Equals,
              value: true
            }
          ]
        }
      ]
    },
    onComplete: {
      type: TransitionType.Branch,
      branches: [
        {
          name: 'Location Change',
          conditions: [
            {
              logic: 'AND',
              conditions: [
                {
                  id: 'branch-location',
                  field: 'process.isLocationChange',
                  operator: ConditionOperator.Equals,
                  value: true
                }
              ]
            }
          ],
          targetStepId: 'STEP-FACILITIES-TASKS'
        },
        {
          name: 'Same Location',
          conditions: [],
          targetStepId: 'STEP-NEW-ROLE-TASKS',
          isDefault: true
        }
      ]
    }
  },

  // Step 8: Facilities Tasks (conditional)
  {
    id: 'STEP-FACILITIES-TASKS',
    name: 'Facilities Setup',
    description: 'Workspace setup at new location',
    type: StepType.AssignTasks,
    order: 8,
    config: {
      taskTitle: 'Prepare workspace for {{employeeName}} at {{newLocation}}',
      assigneeRole: 'Facilities',
      dueDaysFromNow: 5
    },
    onComplete: {
      type: TransitionType.Goto,
      targetStepId: 'STEP-NEW-ROLE-TASKS'
    },
    sla: {
      warningHours: 72,
      breachHours: 120
    }
  },

  // Step 9: New Role Tasks
  {
    id: 'STEP-NEW-ROLE-TASKS',
    name: 'New Role Orientation',
    description: 'Orientation tasks for the new role',
    type: StepType.AssignTasks,
    order: 9,
    config: {
      taskTitle: 'Complete new role orientation',
      assigneeField: 'employeeId',
      dueDaysFromNow: 14
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 240,
      breachHours: 336
    }
  },

  // Step 10: New Manager Introduction
  {
    id: 'STEP-NEW-MGR-INTRO',
    name: 'New Manager Introduction',
    description: 'New manager conducts introduction meeting',
    type: StepType.CreateTask,
    order: 10,
    config: {
      taskTitle: 'Conduct introduction meeting with {{employeeName}}',
      assigneeField: 'newManagerId',
      dueDaysFromNow: 3
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 11: Wait for Orientation Tasks
  {
    id: 'STEP-WAIT-ORIENTATION',
    name: 'Wait for Orientation',
    description: 'Wait for new role orientation to complete',
    type: StepType.WaitForTasks,
    order: 11,
    config: {
      waitForTaskIds: ['STEP-NEW-ROLE-TASKS'],
      waitCondition: 'all'
    },
    onComplete: { type: TransitionType.Next },
    timeoutHours: 480 // 20 days
  },

  // Step 12: 30-Day Review
  {
    id: 'STEP-30DAY-REVIEW',
    name: '30-Day Transfer Review',
    description: 'New manager conducts 30-day check-in',
    type: StepType.CreateTask,
    order: 12,
    config: {
      taskTitle: 'Conduct 30-day transfer review for {{employeeName}}',
      assigneeField: 'newManagerId',
      dueDaysFromNow: 35
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 13: HR Update Records
  {
    id: 'STEP-HR-UPDATE',
    name: 'HR Records Update',
    description: 'HR updates employee records',
    type: StepType.CreateTask,
    order: 13,
    config: {
      taskTitle: 'Update HR records for {{employeeName}} transfer',
      assigneeRole: 'HR Admin',
      dueDaysFromNow: 5
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 14: Completion Notification
  {
    id: 'STEP-COMPLETION-NOTIFY',
    name: 'Transfer Complete Notification',
    description: 'Notify all stakeholders of successful transfer',
    type: StepType.Notification,
    order: 14,
    config: {
      notificationType: 'WorkflowCompleted',
      recipientField: 'employeeEmail',
      notificationSubject: 'Transfer Complete - {{employeeName}}',
      messageTemplate: 'Your transfer to {{newDepartment}} has been successfully completed. Welcome to your new role!'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 15: End
  {
    id: 'STEP-END',
    name: 'End Transfer',
    description: 'Complete the transfer workflow',
    type: StepType.End,
    order: 15,
    config: {},
    onComplete: { type: TransitionType.End }
  }
];

export const MOVER_WORKFLOW_VARIABLES: IWorkflowVariable[] = [
  { name: 'employeeName', type: 'string', description: 'Name of the transferring employee' },
  { name: 'employeeEmail', type: 'string', description: 'Email of the employee' },
  { name: 'currentDepartment', type: 'string', description: 'Current department' },
  { name: 'newDepartment', type: 'string', description: 'New department' },
  { name: 'currentManagerId', type: 'number', description: 'Current manager SharePoint user ID' },
  { name: 'newManagerId', type: 'number', description: 'New manager SharePoint user ID' },
  { name: 'transferDate', type: 'date', description: 'Effective transfer date' },
  { name: 'isLocationChange', type: 'boolean', defaultValue: false, description: 'Whether physical relocation is required' },
  { name: 'currentLocation', type: 'string', description: 'Current office location' },
  { name: 'newLocation', type: 'string', description: 'New office location' }
];

export const DEFAULT_MOVER_WORKFLOW = {
  Title: 'Standard Transfer/Relocation Workflow',
  WorkflowCode: 'WF-MOVER-STD-001',
  Description: 'Standard workflow for employee transfers and relocations. Handles knowledge transfer, access updates, and new role orientation.',
  Version: '1.0.0',
  ProcessType: ProcessType.Mover,
  IsActive: true,
  IsDefault: true,
  Category: 'Transfer',
  Tags: 'transfer,mover,relocation,role-change',
  EstimatedDuration: 720, // 30 days in hours
  Steps: JSON.stringify(MOVER_WORKFLOW_STEPS),
  Variables: JSON.stringify(MOVER_WORKFLOW_VARIABLES),
  TriggerConditions: JSON.stringify([])
};

// ============================================================================
// LEAVER (OFFBOARDING) WORKFLOW DEFINITION
// ============================================================================

export const LEAVER_WORKFLOW_STEPS: IWorkflowStep[] = [
  // Step 1: Start
  {
    id: 'STEP-START',
    name: 'Start Offboarding',
    description: 'Initialize the offboarding workflow',
    type: StepType.Start,
    order: 1,
    config: {},
    onComplete: { type: TransitionType.Next }
  },

  // Step 2: Notify Manager
  {
    id: 'STEP-NOTIFY-MANAGER',
    name: 'Notify Manager',
    description: 'Notify manager about employee departure',
    type: StepType.Notification,
    order: 2,
    config: {
      notificationType: 'ProcessStarted',
      recipientField: 'managerId',
      notificationSubject: 'Employee Departure: {{employeeName}} - Action Required',
      messageTemplate: '{{employeeName}} will be leaving on {{lastWorkingDay}}. Please initiate knowledge transfer and complete manager offboarding tasks urgently.'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 3: Notify HR
  {
    id: 'STEP-NOTIFY-HR',
    name: 'Notify HR',
    description: 'Notify HR team about the departure',
    type: StepType.Notification,
    order: 3,
    config: {
      notificationType: 'ProcessStarted',
      recipientRole: 'HR Admin',
      notificationSubject: 'Employee Departure: {{employeeName}} - HR Action Required',
      messageTemplate: 'Offboarding process initiated for {{employeeName}}. Department: {{department}}. Last working day: {{lastWorkingDay}}. Please complete HR offboarding tasks.'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 4: Knowledge Transfer Tasks
  {
    id: 'STEP-KNOWLEDGE-TRANSFER',
    name: 'Knowledge Transfer',
    description: 'Employee completes knowledge transfer documentation',
    type: StepType.AssignTasks,
    order: 4,
    config: {
      taskTitle: 'Complete knowledge transfer and handover documentation',
      assigneeField: 'employeeId',
      dueDaysFromNow: 7
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 96,
      breachHours: 144,
      escalateTo: 'managerId'
    }
  },

  // Step 5: Manager Handover Review
  {
    id: 'STEP-HANDOVER-APPROVAL',
    name: 'Handover Review',
    description: 'Manager reviews and approves knowledge transfer',
    type: StepType.Approval,
    order: 5,
    config: {
      approverField: 'managerId'
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 24,
      breachHours: 48
    }
  },

  // Step 6: IT Access Revocation
  {
    id: 'STEP-IT-ACCESS-REVOKE',
    name: 'IT Access Revocation',
    description: 'Revoke all IT access and system permissions',
    type: StepType.AssignTasks,
    order: 6,
    config: {
      taskTitle: 'Revoke IT access for {{employeeName}}',
      assigneeRole: 'IT Admin',
      dueDaysFromNow: 1 // Day before or on last day
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 8,
      breachHours: 24
    }
  },

  // Step 7: Asset Return
  {
    id: 'STEP-ASSET-RETURN',
    name: 'Asset Return',
    description: 'Employee returns all company assets',
    type: StepType.AssignTasks,
    order: 7,
    config: {
      taskTitle: 'Return all company assets (laptop, badge, phone, etc.)',
      assigneeField: 'employeeId',
      dueDaysFromNow: 0 // On last working day
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 4,
      breachHours: 8
    }
  },

  // Step 8: Asset Verification
  {
    id: 'STEP-ASSET-VERIFY',
    name: 'Asset Verification',
    description: 'IT verifies all assets returned',
    type: StepType.Approval,
    order: 8,
    config: {
      approverRole: 'IT Admin'
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 8,
      breachHours: 24
    }
  },

  // Step 9: Exit Interview
  {
    id: 'STEP-EXIT-INTERVIEW',
    name: 'Exit Interview',
    description: 'HR conducts exit interview',
    type: StepType.CreateTask,
    order: 9,
    config: {
      taskTitle: 'Conduct exit interview with {{employeeName}}',
      assigneeRole: 'HR Admin',
      dueDaysFromNow: 0
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 10: Final Payroll Processing
  {
    id: 'STEP-FINAL-PAYROLL',
    name: 'Final Payroll',
    description: 'Process final payroll and benefits',
    type: StepType.CreateTask,
    order: 10,
    config: {
      taskTitle: 'Process final payroll for {{employeeName}}',
      assigneeRole: 'Finance Admin',
      dueDaysFromNow: 3
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 48,
      breachHours: 72
    }
  },

  // Step 11: Badge/Access Deactivation
  {
    id: 'STEP-BADGE-DEACTIVATE',
    name: 'Badge Deactivation',
    description: 'Deactivate building access badge',
    type: StepType.CreateTask,
    order: 11,
    config: {
      taskTitle: 'Deactivate building access for {{employeeName}}',
      assigneeRole: 'Facilities',
      dueDaysFromNow: 0
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 12: HR Final Documentation
  {
    id: 'STEP-HR-FINAL-DOCS',
    name: 'HR Final Documentation',
    description: 'Complete all HR offboarding documentation',
    type: StepType.AssignTasks,
    order: 12,
    config: {
      taskTitle: 'Complete offboarding documentation for {{employeeName}}',
      assigneeRole: 'HR Admin',
      dueDaysFromNow: 5
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 13: Wait for All Tasks
  {
    id: 'STEP-WAIT-ALL',
    name: 'Wait for Completion',
    description: 'Wait for all offboarding tasks to complete',
    type: StepType.WaitForTasks,
    order: 13,
    config: {
      waitForTaskIds: [
        'STEP-KNOWLEDGE-TRANSFER',
        'STEP-IT-ACCESS-REVOKE',
        'STEP-ASSET-RETURN',
        'STEP-HR-FINAL-DOCS'
      ],
      waitCondition: 'all'
    },
    onComplete: { type: TransitionType.Next },
    timeoutHours: 336 // 14 days
  },

  // Step 14: Final Verification
  {
    id: 'STEP-FINAL-VERIFY',
    name: 'Final Verification',
    description: 'HR verifies all offboarding complete',
    type: StepType.Approval,
    order: 14,
    config: {
      approverRole: 'HR Manager'
    },
    onComplete: { type: TransitionType.Next },
    sla: {
      warningHours: 24,
      breachHours: 48
    }
  },

  // Step 15: Departure Notification
  {
    id: 'STEP-DEPARTURE-NOTIFY',
    name: 'Departure Notification',
    description: 'Send farewell notification',
    type: StepType.Notification,
    order: 15,
    config: {
      notificationType: 'WorkflowCompleted',
      recipientField: 'employeeEmail',
      notificationSubject: 'Farewell - {{employeeName}}',
      messageTemplate: 'Your offboarding is now complete. We wish you all the best in your future endeavors. Thank you for your contributions to the team.'
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 16: Archive Employee Record
  {
    id: 'STEP-ARCHIVE-RECORD',
    name: 'Archive Record',
    description: 'Archive employee record',
    type: StepType.Action,
    order: 16,
    config: {
      actionType: ActionType.UpdateListItem,
      actionConfig: {
        listName: 'JML_Employees',
        itemIdField: 'employeeId',
        updates: [
          { fieldName: 'Status', value: 'Archived' },
          { fieldName: 'ArchiveDate', valueField: '{{today}}' }
        ]
      }
    },
    onComplete: { type: TransitionType.Next }
  },

  // Step 17: End
  {
    id: 'STEP-END',
    name: 'End Offboarding',
    description: 'Complete the offboarding workflow',
    type: StepType.End,
    order: 17,
    config: {},
    onComplete: { type: TransitionType.End }
  }
];

export const LEAVER_WORKFLOW_VARIABLES: IWorkflowVariable[] = [
  { name: 'employeeName', type: 'string', description: 'Name of the departing employee' },
  { name: 'employeeEmail', type: 'string', description: 'Email of the employee' },
  { name: 'employeeId', type: 'number', description: 'Employee record ID' },
  { name: 'department', type: 'string', description: 'Employee department' },
  { name: 'managerId', type: 'number', description: 'Manager SharePoint user ID' },
  { name: 'lastWorkingDay', type: 'date', description: 'Last working day' },
  { name: 'resignationType', type: 'string', description: 'Voluntary/Involuntary/Retirement' },
  { name: 'exitInterviewCompleted', type: 'boolean', defaultValue: false, description: 'Exit interview status' },
  { name: 'assetsReturned', type: 'boolean', defaultValue: false, description: 'Asset return status' }
];

export const DEFAULT_LEAVER_WORKFLOW = {
  Title: 'Standard Offboarding Workflow',
  WorkflowCode: 'WF-LEAVER-STD-001',
  Description: 'Standard workflow for employee offboarding. Covers knowledge transfer, access revocation, asset return, and final documentation.',
  Version: '1.0.0',
  ProcessType: ProcessType.Leaver,
  IsActive: true,
  IsDefault: true,
  Category: 'Offboarding',
  Tags: 'offboarding,leaver,departure,exit',
  EstimatedDuration: 336, // 14 days in hours
  Steps: JSON.stringify(LEAVER_WORKFLOW_STEPS),
  Variables: JSON.stringify(LEAVER_WORKFLOW_VARIABLES),
  TriggerConditions: JSON.stringify([])
};

// ============================================================================
// EXPORT ALL DEFAULTS
// ============================================================================

export const ALL_DEFAULT_WORKFLOWS = [
  DEFAULT_JOINER_WORKFLOW,
  DEFAULT_MOVER_WORKFLOW,
  DEFAULT_LEAVER_WORKFLOW
];

/**
 * Get default workflow definition by process type
 */
export function getDefaultWorkflowForProcessType(processType: ProcessType): typeof DEFAULT_JOINER_WORKFLOW | typeof DEFAULT_MOVER_WORKFLOW | typeof DEFAULT_LEAVER_WORKFLOW | undefined {
  switch (processType) {
    case ProcessType.Joiner:
      return DEFAULT_JOINER_WORKFLOW;
    case ProcessType.Mover:
      return DEFAULT_MOVER_WORKFLOW;
    case ProcessType.Leaver:
      return DEFAULT_LEAVER_WORKFLOW;
    default:
      return undefined;
  }
}
