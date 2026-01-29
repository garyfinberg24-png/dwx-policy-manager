// PM_TemplateTaskMapping List Model

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlTemplateTaskMapping extends IBaseListItem {
  // References
  TemplateIDId?: number; // Lookup to PM_ChecklistTemplates
  TaskIDId?: number; // Lookup to PM_Tasks

  // Sequencing
  SequenceOrder: number; // Order in which tasks should be executed
  IsMandatory: boolean;

  // Timing
  OffsetDays: number; // Days offset from process start date

  // Dependencies
  DependsOnTaskID?: string; // TaskCode of prerequisite task
  CustomInstructions?: string; // Override default task instructions

  // Overrides
  OverrideAssigneeId?: number;
  OverrideAssignee?: IUser;
  OverrideSLAHours?: number;

  // Status
  IsActive: boolean;
}

// For building checklists from templates
export interface ITemplateTaskMappingExpanded {
  Id: number;
  TemplateId: number;
  TemplateName: string;
  TaskId: number;
  TaskTitle: string;
  TaskCode: string;
  TaskCategory: string;
  SequenceOrder: number;
  IsMandatory: boolean;
  OffsetDays: number;
  DependsOnTaskID?: string;
  EstimatedHours?: number;
  SLAHours?: number;
  DefaultAssignee?: IUser;
}

// For wizard/process creation
export interface ITemplateWithTasks {
  templateId: number;
  templateName: string;
  templateCode: string;
  processType: string;
  tasks: ITemplateTaskMappingExpanded[];
  totalTasks: number;
  estimatedDuration: number;
}
