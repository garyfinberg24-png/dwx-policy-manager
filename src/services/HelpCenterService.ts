// @ts-nocheck
// Help Center Service - Manages help articles, cheatsheets, and tickets

import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { IHelpArticle, IHelpArticleSearchResult, HelpArticleCategory, ArticleType } from "../models/IHelpArticle";
import { ICheatsheet, CheatsheetCategory, ICheatsheetItem } from "../models/ICheatsheet";
import { IHelpTicket, TicketCategory, TicketStatus, SeverityLevel } from "../models/IHelpTicket";
import { logger } from "./LoggingService";

export class HelpCenterService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  // ============================================
  // Help Articles
  // ============================================

  /**
   * Get all published help articles
   */
  public async getArticles(category?: HelpArticleCategory): Promise<IHelpArticle[]> {
    try {
      let query = this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .select(
          "Id", "Title", "Category", "ArticleType", "Summary", "Content",
          "Keywords", "ThumbnailUrl", "VideoUrl", "ParentArticleId", "SortOrder",
          "IsPublished", "IsFeatured", "ViewCount", "HelpfulCount", "NotHelpfulCount",
          "LastReviewedDate", "RelatedArticles", "RelatedWebParts",
          "VersionNumber", "ChangeLog", "Created", "Modified",
          "Author/Id", "Author/Title", "Author/EMail",
          "ReviewedBy/Id", "ReviewedBy/Title", "ReviewedBy/EMail"
        )
        .expand("Author", "ReviewedBy")
        .filter("IsPublished eq 1")
        .orderBy("SortOrder", true)
        .orderBy("Title", true)
        .top(500);

      if (category) {
        query = query.filter(`Category eq '${category}'`);
      }

      const items = await query();

      return items.map(item => this._mapHelpArticle(item));
    } catch (error) {
      logger.error("HelpCenterService", "Error fetching articles", error);
      throw new Error(`Failed to fetch help articles: ${error.message}`);
    }
  }

  /**
   * Get featured articles for home page
   */
  public async getFeaturedArticles(): Promise<IHelpArticle[]> {
    try {
      const items = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .select(
          "Id", "Title", "Category", "ArticleType", "Summary", "Content",
          "Keywords", "ThumbnailUrl", "ViewCount", "HelpfulCount", "NotHelpfulCount",
          "IsFeatured", "IsPublished"
        )
        .filter("IsPublished eq 1 and IsFeatured eq 1")
        .orderBy("SortOrder", true)
        .top(6)();

      return items.map(item => this._mapHelpArticle(item));
    } catch (error) {
      logger.error("HelpCenterService", "Error fetching featured articles", error);
      throw new Error(`Failed to fetch featured articles: ${error.message}`);
    }
  }

  /**
   * Get most popular articles
   */
  public async getPopularArticles(limit: number = 10): Promise<IHelpArticle[]> {
    try {
      const items = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .select(
          "Id", "Title", "Category", "ArticleType", "Summary",
          "ViewCount", "HelpfulCount", "ThumbnailUrl"
        )
        .filter("IsPublished eq 1")
        .orderBy("ViewCount", false)
        .top(limit)();

      return items.map(item => this._mapHelpArticle(item));
    } catch (error) {
      logger.error("HelpCenterService", "Error fetching popular articles", error);
      throw new Error(`Failed to fetch popular articles: ${error.message}`);
    }
  }

  /**
   * Get article by ID
   */
  public async getArticleById(articleId: number): Promise<IHelpArticle> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .getById(articleId)
        .select(
          "Id", "Title", "Category", "ArticleType", "Summary", "Content",
          "Keywords", "ThumbnailUrl", "VideoUrl", "ParentArticleId", "SortOrder",
          "IsPublished", "IsFeatured", "ViewCount", "HelpfulCount", "NotHelpfulCount",
          "LastReviewedDate", "RelatedArticles", "RelatedWebParts",
          "VersionNumber", "ChangeLog", "Created", "Modified",
          "Author/Id", "Author/Title", "Author/EMail",
          "ReviewedBy/Id", "ReviewedBy/Title", "ReviewedBy/EMail"
        )
        .expand("Author", "ReviewedBy")();

      // Increment view count
      await this.incrementArticleViewCount(articleId);

      return this._mapHelpArticle(item);
    } catch (error) {
      logger.error("HelpCenterService", `Error fetching article ${articleId}`, error);
      throw new Error(`Failed to fetch article: ${error.message}`);
    }
  }

  /**
   * Search articles by keyword
   */
  public async searchArticles(searchTerm: string, category?: HelpArticleCategory, articleType?: ArticleType): Promise<IHelpArticleSearchResult[]> {
    try {
      // Build search filter
      const searchFilter = `(substringof('${searchTerm}', Title) or substringof('${searchTerm}', Summary) or substringof('${searchTerm}', Content) or substringof('${searchTerm}', Keywords))`;
      let filter = `IsPublished eq 1 and ${searchFilter}`;

      if (category) {
        filter += ` and Category eq '${category}'`;
      }

      if (articleType) {
        filter += ` and ArticleType eq '${articleType}'`;
      }

      const items = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .select(
          "Id", "Title", "Category", "ArticleType", "Summary", "Content",
          "Keywords", "ThumbnailUrl", "ViewCount", "HelpfulCount"
        )
        .filter(filter)
        .orderBy("ViewCount", false)
        .top(50)();

      // Calculate search relevance score
      return items.map(item => {
        const article = this._mapHelpArticle(item);
        const score = this._calculateSearchScore(article, searchTerm);
        const matchedKeywords = this._getMatchedKeywords(article, searchTerm);

        return {
          ...article,
          SearchScore: score,
          MatchedKeywords: matchedKeywords
        } as IHelpArticleSearchResult;
      }).sort((a, b) => (b.SearchScore || 0) - (a.SearchScore || 0));
    } catch (error) {
      logger.error("HelpCenterService", "Error searching articles", error);
      throw new Error(`Failed to search articles: ${error.message}`);
    }
  }

  /**
   * Increment article view count
   */
  public async incrementArticleViewCount(articleId: number): Promise<void> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .getById(articleId)
        .select("ViewCount")();

      const currentCount = item.ViewCount || 0;

      await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .getById(articleId)
        .update({
          ViewCount: currentCount + 1
        });
    } catch (error) {
      logger.warn("HelpCenterService", `Error incrementing view count for article ${articleId}`, error);
      // Don't throw - view count is not critical
    }
  }

  /**
   * Submit article feedback (helpful/not helpful)
   */
  public async submitArticleFeedback(articleId: number, isHelpful: boolean): Promise<void> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .getById(articleId)
        .select("HelpfulCount", "NotHelpfulCount")();

      const update: any = {};

      if (isHelpful) {
        update.HelpfulCount = (item.HelpfulCount || 0) + 1;
      } else {
        update.NotHelpfulCount = (item.NotHelpfulCount || 0) + 1;
      }

      await this._sp.web.lists.getByTitle("PM_HelpArticles").items
        .getById(articleId)
        .update(update);

      logger.info("HelpCenterService", `Feedback submitted for article ${articleId}: ${isHelpful ? 'helpful' : 'not helpful'}`);
    } catch (error) {
      logger.error("HelpCenterService", `Error submitting feedback for article ${articleId}`, error);
      throw new Error(`Failed to submit feedback: ${error.message}`);
    }
  }

  // ============================================
  // Cheatsheets
  // ============================================

  /**
   * Get all published cheatsheets
   */
  public async getCheatsheets(category?: CheatsheetCategory): Promise<ICheatsheet[]> {
    try {
      let query = this._sp.web.lists.getByTitle("PM_Cheatsheets").items
        .select(
          "Id", "Title", "Category", "Description", "ItemsJSON",
          "DisplayFormat", "IconName", "ColorTheme", "SortOrder",
          "IsPublished", "IsPinned", "LastReviewedDate", "ViewCount",
          "RelatedArticles", "RelatedWebParts", "Tags",
          "Author/Id", "Author/Title", "Author/EMail"
        )
        .expand("Author")
        .filter("IsPublished eq 1")
        .orderBy("IsPinned", false)
        .orderBy("SortOrder", true)
        .top(100);

      if (category) {
        query = query.filter(`Category eq '${category}'`);
      }

      const items = await query();

      return items.map(item => this._mapCheatsheet(item));
    } catch (error) {
      logger.error("HelpCenterService", "Error fetching cheatsheets", error);
      throw new Error(`Failed to fetch cheatsheets: ${error.message}`);
    }
  }

  /**
   * Get cheatsheet by ID
   */
  public async getCheatsheetById(cheatsheetId: number): Promise<ICheatsheet> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_Cheatsheets").items
        .getById(cheatsheetId)
        .select(
          "Id", "Title", "Category", "Description", "ItemsJSON",
          "DisplayFormat", "IconName", "ColorTheme", "SortOrder",
          "IsPublished", "IsPinned", "LastReviewedDate", "ViewCount",
          "RelatedArticles", "RelatedWebParts", "Tags",
          "Author/Id", "Author/Title", "Author/EMail"
        )
        .expand("Author")();

      // Increment view count
      await this.incrementCheatsheetViewCount(cheatsheetId);

      return this._mapCheatsheet(item);
    } catch (error) {
      logger.error("HelpCenterService", `Error fetching cheatsheet ${cheatsheetId}`, error);
      throw new Error(`Failed to fetch cheatsheet: ${error.message}`);
    }
  }

  /**
   * Increment cheatsheet view count
   */
  public async incrementCheatsheetViewCount(cheatsheetId: number): Promise<void> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_Cheatsheets").items
        .getById(cheatsheetId)
        .select("ViewCount")();

      const currentCount = item.ViewCount || 0;

      await this._sp.web.lists.getByTitle("PM_Cheatsheets").items
        .getById(cheatsheetId)
        .update({
          ViewCount: currentCount + 1
        });
    } catch (error) {
      logger.warn("HelpCenterService", `Error incrementing view count for cheatsheet ${cheatsheetId}`, error);
      // Don't throw - view count is not critical
    }
  }

  // ============================================
  // Help Tickets
  // ============================================

  /**
   * Create a new help ticket
   */
  public async createTicket(ticket: Partial<IHelpTicket>): Promise<number> {
    try {
      // Generate ticket number
      const ticketNumber = await this._generateTicketNumber();

      const ticketData: any = {
        Title: ticket.Title,
        TicketNumber: ticketNumber,
        Category: ticket.Category,
        Status: "New",
        Priority: ticket.Priority || "Normal",
        Severity: ticket.Severity || "Medium",
        Description: ticket.Description,
        StepsToReproduce: ticket.StepsToReproduce,
        ExpectedBehavior: ticket.ExpectedBehavior,
        ActualBehavior: ticket.ActualBehavior,
        WebPartName: ticket.WebPartName,
        PageUrl: ticket.PageUrl,
        BrowserInfo: ticket.BrowserInfo,
        ErrorMessage: ticket.ErrorMessage,
        ScreenshotUrl: ticket.ScreenshotUrl,
        SubmittedDate: new Date().toISOString(),
        RequiresFollowUp: false
      };

      // Add submitted by if provided
      if (ticket.SubmittedById) {
        ticketData.SubmittedById = ticket.SubmittedById;
      }

      const result = await this._sp.web.lists.getByTitle("PM_HelpTickets").items.add(ticketData);

      logger.info("HelpCenterService", `Ticket created: ${ticketNumber} (ID: ${result.data.Id})`);

      return result.data.Id;
    } catch (error) {
      logger.error("HelpCenterService", "Error creating ticket", error);
      throw new Error(`Failed to create ticket: ${error.message}`);
    }
  }

  /**
   * Get tickets for current user
   */
  public async getMyTickets(userId: number): Promise<IHelpTicket[]> {
    try {
      const items = await this._sp.web.lists.getByTitle("PM_HelpTickets").items
        .select(
          "Id", "Title", "TicketNumber", "Category", "Status", "Priority", "Severity",
          "Description", "SubmittedDate", "ResolvedDate", "SatisfactionRating",
          "SubmittedBy/Id", "SubmittedBy/Title", "SubmittedBy/EMail",
          "AssignedTo/Id", "AssignedTo/Title", "AssignedTo/EMail"
        )
        .expand("SubmittedBy", "AssignedTo")
        .filter(`SubmittedById eq ${userId}`)
        .orderBy("SubmittedDate", false)
        .top(100)();

      return items.map(item => this._mapHelpTicket(item));
    } catch (error) {
      logger.error("HelpCenterService", "Error fetching my tickets", error);
      throw new Error(`Failed to fetch tickets: ${error.message}`);
    }
  }

  /**
   * Get ticket by ID
   */
  public async getTicketById(ticketId: number): Promise<IHelpTicket> {
    try {
      const item = await this._sp.web.lists.getByTitle("PM_HelpTickets").items
        .getById(ticketId)
        .select(
          "Id", "Title", "TicketNumber", "Category", "Status", "Priority", "Severity",
          "Description", "StepsToReproduce", "ExpectedBehavior", "ActualBehavior",
          "WebPartName", "PageUrl", "BrowserInfo", "ErrorMessage", "ScreenshotUrl",
          "SubmittedDate", "FirstResponseDate", "ResolvedDate", "ClosedDate",
          "Resolution", "ResolutionNotes", "RelatedArticleId",
          "ResponseTime", "ResolutionTime", "SatisfactionRating", "FeedbackComments",
          "RelatedTickets", "RelatedProcessId", "RequiresFollowUp", "FollowUpDate", "FollowUpNotes",
          "SubmittedBy/Id", "SubmittedBy/Title", "SubmittedBy/EMail",
          "AssignedTo/Id", "AssignedTo/Title", "AssignedTo/EMail"
        )
        .expand("SubmittedBy", "AssignedTo")();

      return this._mapHelpTicket(item);
    } catch (error) {
      logger.error("HelpCenterService", `Error fetching ticket ${ticketId}`, error);
      throw new Error(`Failed to fetch ticket: ${error.message}`);
    }
  }

  // ============================================
  // Private Helper Methods
  // ============================================

  private _mapHelpArticle(item: any): IHelpArticle {
    return {
      Id: item.Id,
      Title: item.Title,
      Category: item.Category,
      ArticleType: item.ArticleType,
      Summary: item.Summary,
      Content: item.Content,
      Keywords: item.Keywords,
      ThumbnailUrl: item.ThumbnailUrl,
      VideoUrl: item.VideoUrl,
      ParentArticleId: item.ParentArticleId,
      SortOrder: item.SortOrder,
      IsPublished: item.IsPublished,
      IsFeatured: item.IsFeatured,
      ViewCount: item.ViewCount,
      HelpfulCount: item.HelpfulCount,
      NotHelpfulCount: item.NotHelpfulCount,
      LastReviewedDate: item.LastReviewedDate ? new Date(item.LastReviewedDate) : undefined,
      RelatedArticles: item.RelatedArticles,
      RelatedWebParts: item.RelatedWebParts,
      VersionNumber: item.VersionNumber,
      ChangeLog: item.ChangeLog,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      AuthorId: item.Author?.Id,
      Author: item.Author ? {
        Id: item.Author.Id,
        Title: item.Author.Title,
        EMail: item.Author.EMail
      } : undefined,
      ReviewedById: item.ReviewedBy?.Id,
      ReviewedBy: item.ReviewedBy ? {
        Id: item.ReviewedBy.Id,
        Title: item.ReviewedBy.Title,
        EMail: item.ReviewedBy.EMail
      } : undefined
    };
  }

  private _mapCheatsheet(item: any): ICheatsheet {
    // Parse JSON items
    let items: ICheatsheetItem[] = [];
    if (item.ItemsJSON) {
      try {
        items = JSON.parse(item.ItemsJSON);
      } catch (e) {
        logger.warn("HelpCenterService", `Failed to parse ItemsJSON for cheatsheet ${item.Id}`, e);
      }
    }

    return {
      Id: item.Id,
      Title: item.Title,
      Category: item.Category,
      Description: item.Description,
      Items: items,
      DisplayFormat: item.DisplayFormat,
      IconName: item.IconName,
      ColorTheme: item.ColorTheme,
      SortOrder: item.SortOrder,
      IsPublished: item.IsPublished,
      IsPinned: item.IsPinned,
      LastReviewedDate: item.LastReviewedDate ? new Date(item.LastReviewedDate) : undefined,
      ViewCount: item.ViewCount,
      RelatedArticles: item.RelatedArticles,
      RelatedWebParts: item.RelatedWebParts,
      Tags: item.Tags,
      AuthorId: item.Author?.Id,
      Author: item.Author ? {
        Id: item.Author.Id,
        Title: item.Author.Title,
        EMail: item.Author.EMail
      } : undefined
    };
  }

  private _mapHelpTicket(item: any): IHelpTicket {
    return {
      Id: item.Id,
      Title: item.Title,
      TicketNumber: item.TicketNumber,
      Category: item.Category,
      Status: item.Status,
      Priority: item.Priority,
      Severity: item.Severity,
      Description: item.Description,
      StepsToReproduce: item.StepsToReproduce,
      ExpectedBehavior: item.ExpectedBehavior,
      ActualBehavior: item.ActualBehavior,
      WebPartName: item.WebPartName,
      PageUrl: item.PageUrl,
      BrowserInfo: item.BrowserInfo,
      ErrorMessage: item.ErrorMessage,
      ScreenshotUrl: item.ScreenshotUrl,
      SubmittedDate: item.SubmittedDate ? new Date(item.SubmittedDate) : undefined,
      FirstResponseDate: item.FirstResponseDate ? new Date(item.FirstResponseDate) : undefined,
      ResolvedDate: item.ResolvedDate ? new Date(item.ResolvedDate) : undefined,
      ClosedDate: item.ClosedDate ? new Date(item.ClosedDate) : undefined,
      Resolution: item.Resolution,
      ResolutionNotes: item.ResolutionNotes,
      RelatedArticleId: item.RelatedArticleId,
      ResponseTime: item.ResponseTime,
      ResolutionTime: item.ResolutionTime,
      SatisfactionRating: item.SatisfactionRating,
      FeedbackComments: item.FeedbackComments,
      RelatedTickets: item.RelatedTickets,
      RelatedProcessId: item.RelatedProcessId,
      RequiresFollowUp: item.RequiresFollowUp,
      FollowUpDate: item.FollowUpDate ? new Date(item.FollowUpDate) : undefined,
      FollowUpNotes: item.FollowUpNotes,
      SubmittedById: item.SubmittedBy?.Id,
      SubmittedBy: item.SubmittedBy ? {
        Id: item.SubmittedBy.Id,
        Title: item.SubmittedBy.Title,
        EMail: item.SubmittedBy.EMail
      } : undefined,
      AssignedToId: item.AssignedTo?.Id,
      AssignedTo: item.AssignedTo ? {
        Id: item.AssignedTo.Id,
        Title: item.AssignedTo.Title,
        EMail: item.AssignedTo.EMail
      } : undefined
    };
  }

  private _calculateSearchScore(article: IHelpArticle, searchTerm: string): number {
    const term = searchTerm.toLowerCase();
    let score = 0;

    // Title match (highest weight)
    if (article.Title.toLowerCase().includes(term)) {
      score += 100;
    }

    // Keywords match
    if (article.Keywords && article.Keywords.toLowerCase().includes(term)) {
      score += 50;
    }

    // Summary match
    if (article.Summary && article.Summary.toLowerCase().includes(term)) {
      score += 30;
    }

    // Content match (lowest weight)
    if (article.Content && article.Content.toLowerCase().includes(term)) {
      score += 10;
    }

    // Boost for featured articles
    if (article.IsFeatured) {
      score += 20;
    }

    // Boost for popular articles (view count)
    if (article.ViewCount && article.ViewCount > 100) {
      score += 10;
    }

    return score;
  }

  private _getMatchedKeywords(article: IHelpArticle, searchTerm: string): string[] {
    if (!article.Keywords) return [];

    const term = searchTerm.toLowerCase();
    const keywords = article.Keywords.split(",").map(k => k.trim());

    return keywords.filter(keyword => keyword.toLowerCase().includes(term));
  }

  private async _generateTicketNumber(): Promise<string> {
    const year = new Date().getFullYear();
    const prefix = `HELP-${year}-`;

    try {
      // Get count of tickets this year
      const items = await this._sp.web.lists.getByTitle("PM_HelpTickets").items
        .select("Id")
        .filter(`startswith(TicketNumber, '${prefix}')`)
        .top(5000)();

      const nextNumber = (items.length + 1).toString();
      const paddedNumber = ("0000" + nextNumber).slice(-4);
      return `${prefix}${paddedNumber}`;
    } catch (error) {
      logger.warn("HelpCenterService", "Error generating ticket number, using timestamp", error);
      // Fallback to timestamp-based number
      return `${prefix}${Date.now()}`;
    }
  }
}
