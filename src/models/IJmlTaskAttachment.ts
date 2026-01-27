// JML_TaskAttachments Model
// Manages file attachments for task assignments

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlTaskAttachment extends IBaseListItem {
  // Task Reference
  TaskAssignmentId: number;
  TaskAssignment?: {
    Id: number;
    Title: string;
  };

  // File Information
  FileName: string;
  FileUrl: string;
  FileSize: number; // Bytes
  FileType: string; // Extension (e.g., ".pdf", ".docx")

  // Upload Information
  UploadedById: number;
  UploadedBy?: IUser;
  UploadedDate: Date;

  // Metadata
  Description?: string;
  Category?: AttachmentCategory;
  IsRequired?: boolean; // Required for task completion

  // Version Control
  VersionLabel?: string;
  IsLatestVersion?: boolean;
  PreviousVersionId?: number;

  // SharePoint Library Fields
  ServerRelativeUrl?: string;
  UniqueId?: string;
}

export enum AttachmentCategory {
  Documentation = 'Documentation',
  Form = 'Form',
  Evidence = 'Evidence',
  Reference = 'Reference',
  Template = 'Template',
  Other = 'Other'
}

// View model for displaying attachments
export interface ITaskAttachmentView extends IJmlTaskAttachment {
  FormattedSize: string; // e.g., "2.5 MB"
  IconName: string; // Fluent UI icon name based on file type
  CanDelete: boolean;
  CanDownload: boolean;
  IsImage: boolean;
  ThumbnailUrl?: string;
}

// Form model for uploading attachments
export interface ITaskAttachmentUpload {
  TaskAssignmentId: number;
  File: File;
  Description?: string;
  Category?: AttachmentCategory;
  IsRequired?: boolean;
}

// Summary model for task attachment count
export interface ITaskAttachmentSummary {
  TaskAssignmentId: number;
  TotalCount: number;
  TotalSize: number;
  FormattedTotalSize: string;
  RequiredCount: number;
  ByCategory: {
    [key in AttachmentCategory]?: number;
  };
  LatestAttachment?: IJmlTaskAttachment;
}
