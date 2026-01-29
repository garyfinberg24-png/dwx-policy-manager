// PM_HelpArticles List Model
// Knowledge base articles, FAQs, and documentation

import { IBaseListItem, IUser } from './ICommon';

export type HelpArticleCategory =
  | 'Getting Started'
  | 'Process Management'
  | 'Task Management'
  | 'User Guide'
  | 'Administrator Guide'
  | 'Troubleshooting'
  | 'FAQ'
  | 'Integration'
  | 'Best Practices'
  | 'Release Notes';

export type ArticleType =
  | 'Documentation'
  | 'FAQ'
  | 'Video Tutorial'
  | 'Quick Guide'
  | 'Cheatsheet'
  | 'Troubleshooting Guide';

export interface IHelpArticle extends IBaseListItem {
  // Article Info
  Title: string; // Article title
  Category: HelpArticleCategory;
  ArticleType: ArticleType;

  // Content
  Summary?: string; // Short summary (100 chars)
  Content: string; // Rich text content (HTML)
  Keywords?: string; // Comma-separated keywords for search

  // Media
  ThumbnailUrl?: string;
  VideoUrl?: string;
  AttachmentsUrl?: string;

  // Organization
  ParentArticleId?: number; // For hierarchical articles
  SortOrder?: number;

  // Metadata
  IsPublished: boolean;
  IsFeatured: boolean;
  ViewCount?: number;
  HelpfulCount?: number;
  NotHelpfulCount?: number;

  // Authoring
  AuthorId?: number;
  Author?: IUser;
  LastReviewedDate?: Date;
  ReviewedById?: number;
  ReviewedBy?: IUser;

  // Related Items
  RelatedArticles?: string; // Comma-separated IDs
  RelatedWebParts?: string; // Comma-separated web part names

  // Version Control
  VersionNumber?: string;
  ChangeLog?: string;
}

export interface IHelpArticleSearchResult extends IHelpArticle {
  SearchScore?: number;
  MatchedKeywords?: string[];
}

// For article picker/selector
export interface IHelpArticleOption {
  Id: number;
  Title: string;
  Category: HelpArticleCategory;
  ArticleType: ArticleType;
  Summary?: string;
}
