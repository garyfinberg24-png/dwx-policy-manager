// @ts-nocheck
// TaskCommentsService - Handles task comment CRUD operations and threading

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import {
  IJmlTaskComment,
  ITaskCommentView,
  ITaskCommentForm,
  CommentType
} from '../models/IJmlTaskComment';
import { logger } from './LoggingService';

export class TaskCommentsService {
  private sp: SPFI;
  private listTitle = 'PM_TaskComments';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all comments for a task assignment (with threading)
   */
  public async getTaskComments(taskAssignmentId: number, currentUserId: number): Promise<ITaskCommentView[]> {
    try {
      logger.info('TaskCommentsService', `Getting comments for task ${taskAssignmentId}`);

      const items = await this.sp.web.lists.getByTitle(this.listTitle).items
        .filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select(
          'Id', 'Title', 'Comment', 'CommentType', 'IsPinned', 'IsEdited', 'EditedDate',
          'ParentCommentId', 'AttachmentUrls', 'MentionedUserIds',
          'Created', 'Modified',
          'Author/Id', 'Author/Title', 'Author/EMail',
          'Editor/Id', 'Editor/Title'
        )
        .expand('Author', 'Editor')
        .orderBy('Created', true)
        .top(500)();

      // Build threaded comment structure
      const comments: ITaskCommentView[] = items.map(item => this.mapToCommentView(item, currentUserId));

      return this.buildCommentTree(comments);
    } catch (error) {
      logger.error('TaskCommentsService', `Error getting comments for task ${taskAssignmentId}`, error);
      throw error;
    }
  }

  /**
   * Add a new comment
   */
  public async addComment(comment: ITaskCommentForm, authorId: number): Promise<IJmlTaskComment> {
    try {
      logger.info('TaskCommentsService', `Adding comment to task ${comment.TaskAssignmentId}`);

      const itemData: any = {
        TaskAssignmentId: comment.TaskAssignmentId,
        Comment: comment.Comment,
        CommentType: comment.CommentType,
        ParentCommentId: comment.ParentCommentId,
        IsPinned: comment.IsPinned || false,
        AttachmentUrls: comment.AttachmentUrls ? JSON.stringify(comment.AttachmentUrls) : null,
        MentionedUserIds: comment.MentionedUserIds ? JSON.stringify(comment.MentionedUserIds) : null
      };

      const result = await this.sp.web.lists.getByTitle(this.listTitle).items.add(itemData);

      logger.info('TaskCommentsService', `Comment added successfully: ${result.data.Id}`);

      return result.data as IJmlTaskComment;
    } catch (error) {
      logger.error('TaskCommentsService', 'Error adding comment', error);
      throw error;
    }
  }

  /**
   * Update existing comment
   */
  public async updateComment(commentId: number, updates: Partial<ITaskCommentForm>): Promise<void> {
    try {
      logger.info('TaskCommentsService', `Updating comment ${commentId}`);

      const itemData: any = {
        IsEdited: true,
        EditedDate: new Date().toISOString()
      };

      if (updates.Comment) itemData.Comment = updates.Comment;
      if (updates.CommentType) itemData.CommentType = updates.CommentType;
      if (updates.IsPinned !== undefined) itemData.IsPinned = updates.IsPinned;
      if (updates.AttachmentUrls) itemData.AttachmentUrls = JSON.stringify(updates.AttachmentUrls);

      await this.sp.web.lists.getByTitle(this.listTitle).items.getById(commentId).update(itemData);

      logger.info('TaskCommentsService', 'Comment updated successfully');
    } catch (error) {
      logger.error('TaskCommentsService', `Error updating comment ${commentId}`, error);
      throw error;
    }
  }

  /**
   * Delete a comment (and its replies)
   */
  public async deleteComment(commentId: number): Promise<void> {
    try {
      logger.info('TaskCommentsService', `Deleting comment ${commentId}`);

      // Get all child comments
      const children = await this.sp.web.lists.getByTitle(this.listTitle).items
        .filter(`ParentCommentId eq ${commentId}`)
        .select('Id')();

      // Delete children first
      for (const child of children) {
        await this.sp.web.lists.getByTitle(this.listTitle).items.getById(child.Id).delete();
      }

      // Delete parent comment
      await this.sp.web.lists.getByTitle(this.listTitle).items.getById(commentId).delete();

      logger.info('TaskCommentsService', 'Comment and replies deleted successfully');
    } catch (error) {
      logger.error('TaskCommentsService', `Error deleting comment ${commentId}`, error);
      throw error;
    }
  }

  /**
   * Pin/unpin a comment
   */
  public async togglePinComment(commentId: number, isPinned: boolean): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listTitle).items.getById(commentId).update({
        IsPinned: isPinned
      });

      logger.info('TaskCommentsService', `Comment ${commentId} ${isPinned ? 'pinned' : 'unpinned'}`);
    } catch (error) {
      logger.error('TaskCommentsService', `Error toggling pin on comment ${commentId}`, error);
      throw error;
    }
  }

  /**
   * Get comment count for a task
   */
  public async getCommentCount(taskAssignmentId: number): Promise<number> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.listTitle).items
        .filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select('Id')
        .top(1000)();

      return items.length;
    } catch (error) {
      logger.error('TaskCommentsService', `Error getting comment count for task ${taskAssignmentId}`, error);
      return 0;
    }
  }

  /**
   * Map SharePoint item to comment view model
   */
  private mapToCommentView(item: any, currentUserId: number): ITaskCommentView {
    const comment: ITaskCommentView = {
      ...item,
      AttachmentList: item.AttachmentUrls ? JSON.parse(item.AttachmentUrls) : [],
      CanEdit: item.Author?.Id === currentUserId,
      CanDelete: item.Author?.Id === currentUserId,
      Replies: []
    };

    return comment;
  }

  /**
   * Build threaded comment tree
   */
  private buildCommentTree(comments: ITaskCommentView[]): ITaskCommentView[] {
    const commentMap = new Map<number, ITaskCommentView>();
    const rootComments: ITaskCommentView[] = [];

    // First pass: Create map of all comments
    comments.forEach(comment => {
      commentMap.set(comment.Id, comment);
      comment.Replies = [];
    });

    // Second pass: Build tree structure
    comments.forEach(comment => {
      if (comment.ParentCommentId) {
        const parent = commentMap.get(comment.ParentCommentId);
        if (parent) {
          parent.Replies = parent.Replies || [];
          parent.Replies.push(comment);
        }
      } else {
        rootComments.push(comment);
      }
    });

    // Sort: Pinned first, then by date
    return this.sortComments(rootComments);
  }

  /**
   * Sort comments (pinned first, then by date)
   */
  private sortComments(comments: ITaskCommentView[]): ITaskCommentView[] {
    return comments.sort((a, b) => {
      // Pinned comments first
      if (a.IsPinned && !b.IsPinned) return -1;
      if (!a.IsPinned && b.IsPinned) return 1;

      // Then by date (newest first)
      const dateA = new Date(a.Created).getTime();
      const dateB = new Date(b.Created).getTime();
      return dateB - dateA;
    });
  }
}
