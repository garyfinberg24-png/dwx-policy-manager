// JML_TaskComments List Model
// Provides collaboration and discussion threads on task assignments

import { IBaseListItem, IUser } from './ICommon';

export interface IJmlTaskComment extends IBaseListItem {
  // Parent Task Reference
  TaskAssignmentId: number;
  TaskAssignment?: {
    Id: number;
    Title: string;
  };

  // Comment Content
  Comment: string;
  CommentType: CommentType;

  // Author Information
  AuthorId: number;
  Author?: IUser;

  // Threading Support
  ParentCommentId?: number;
  ParentComment?: IJmlTaskComment;

  // Metadata
  IsPinned?: boolean;
  IsEdited?: boolean;
  EditedDate?: Date;

  // Attachments
  AttachmentUrls?: string; // JSON string array

  // Mentions
  MentionedUserIds?: string; // JSON string array of user IDs

  // Reactions (future enhancement)
  Reactions?: string; // JSON object {userId: emoji}
}

export enum CommentType {
  General = 'General',
  Question = 'Question',
  Blocker = 'Blocker',
  Update = 'Update',
  Resolution = 'Resolution'
}

// View model for displaying comments
export interface ITaskCommentView extends IJmlTaskComment {
  Replies?: ITaskCommentView[];
  CanEdit: boolean;
  CanDelete: boolean;
  AttachmentList?: string[];
  MentionedUsers?: IUser[];
}

// Form model for creating/editing comments
export interface ITaskCommentForm {
  TaskAssignmentId: number;
  Comment: string;
  CommentType: CommentType;
  ParentCommentId?: number;
  IsPinned?: boolean;
  AttachmentUrls?: string[];
  MentionedUserIds?: number[];
}
