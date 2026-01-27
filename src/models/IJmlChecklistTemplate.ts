// JML_ChecklistTemplates List Model

import { IBaseListItem, ProcessType } from './ICommon';

export interface IJmlChecklistTemplate extends IBaseListItem {
  // Template Information
  TemplateCode: string; // Unique identifier (e.g., "JOIN-IT-001")
  ProcessType: ProcessType;

  // Template Details
  Description?: string;
  Department?: string;
  JobRole?: string;

  // Metadata
  EstimatedDuration?: number; // In days
  TaskCount?: number;
  IsActive: boolean;
  Version?: string;

  // Usage Stats
  TimesUsed?: number;
  LastUsedDate?: Date;
  AverageCompletionTime?: number; // In days

  // Additional
  Tags?: string;
  CreatedByDepartment?: string;
  ApprovalRequired?: boolean;
}

// For dropdown/selection lists
export interface IJmlChecklistTemplateOption {
  Id: number;
  Title: string;
  TemplateCode: string;
  ProcessType: ProcessType;
  Description?: string;
}
