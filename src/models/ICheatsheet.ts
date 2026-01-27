// JML_Cheatsheets List Model
// Quick reference guides and shortcuts

import { IBaseListItem, IUser } from './ICommon';

export type CheatsheetCategory =
  | 'Keyboard Shortcuts'
  | 'Process Workflows'
  | 'Task Management'
  | 'Admin Tasks'
  | 'Integration Setup'
  | 'Troubleshooting'
  | 'Best Practices'
  | 'Quick Reference';

export interface ICheatsheetItem {
  Title: string;
  Description?: string;
  Shortcut?: string; // e.g., "Ctrl+S"
  Icon?: string; // Fluent UI icon name
  Category?: string;
  SortOrder?: number;
}

export interface ICheatsheet extends IBaseListItem {
  // Cheatsheet Info
  Title: string;
  Category: CheatsheetCategory;
  Description?: string;

  // Content
  Items: ICheatsheetItem[]; // JSON array of cheatsheet items

  // Display
  DisplayFormat: 'Table' | 'Cards' | 'List';
  IconName?: string; // Fluent UI icon
  ColorTheme?: string; // Hex color for card background

  // Organization
  SortOrder?: number;
  IsPublished: boolean;
  IsPinned: boolean; // Show at top

  // Metadata
  AuthorId?: number;
  Author?: IUser;
  LastReviewedDate?: Date;
  ViewCount?: number;

  // Related
  RelatedArticles?: string; // Comma-separated article IDs
  RelatedWebParts?: string; // Comma-separated web part names
  Tags?: string; // Comma-separated tags
}

// For rendering
export interface ICheatsheetRenderData {
  Id: number;
  Title: string;
  Category: CheatsheetCategory;
  Items: ICheatsheetItem[];
  DisplayFormat: 'Table' | 'Cards' | 'List';
  IconName?: string;
  ColorTheme?: string;
}
