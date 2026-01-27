// @ts-nocheck
/**
 * Policy Document Comparison Service
 * Provides version comparison, diff visualization, and change tracking
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IPolicy } from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists, ComparisonLists } from '../constants/SharePointListNames';

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Types of changes detected
 */
export enum ChangeType {
  Added = 'Added',
  Removed = 'Removed',
  Modified = 'Modified',
  Moved = 'Moved',
  Unchanged = 'Unchanged'
}

/**
 * Content block types
 */
export enum BlockType {
  Paragraph = 'Paragraph',
  Heading = 'Heading',
  List = 'List',
  Table = 'Table',
  Image = 'Image',
  Section = 'Section'
}

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Policy version for comparison
 */
export interface IPolicyVersion {
  id: number;
  policyId: number;
  version: string;
  versionNumber: number;
  title: string;
  content: string;
  contentHtml?: string;
  summary?: string;
  effectiveDate?: string;
  createdDate: string;
  createdById: number;
  createdByName: string;
  changeNotes?: string;
  status: string;
  wordCount?: number;
  sectionCount?: number;
}

/**
 * Individual change item
 */
export interface IChangeItem {
  id: string;
  type: ChangeType;
  blockType: BlockType;
  sectionNumber?: string;
  sectionTitle?: string;

  // Original content (for removed/modified)
  originalText?: string;
  originalHtml?: string;
  originalPosition?: number;

  // New content (for added/modified)
  newText?: string;
  newHtml?: string;
  newPosition?: number;

  // For word-level diff
  wordChanges?: IWordChange[];

  // Significance
  significance: 'Major' | 'Minor' | 'Cosmetic';
  category?: string;
}

/**
 * Word-level change
 */
export interface IWordChange {
  type: ChangeType;
  text: string;
  position: number;
}

/**
 * Comparison result
 */
export interface IComparisonResult {
  sourceVersion: IPolicyVersion;
  targetVersion: IPolicyVersion;

  // Summary statistics
  summary: {
    totalChanges: number;
    additions: number;
    deletions: number;
    modifications: number;
    majorChanges: number;
    minorChanges: number;
    cosmeticChanges: number;
    wordCountChange: number;
    percentageChanged: number;
  };

  // Detailed changes
  changes: IChangeItem[];

  // Section-by-section comparison
  sectionComparisons: ISectionComparison[];

  // Metadata
  comparedDate: string;
  comparedById?: number;
  comparedByName?: string;
}

/**
 * Section comparison
 */
export interface ISectionComparison {
  sectionNumber: string;
  sectionTitle: string;
  status: ChangeType;
  originalContent?: string;
  newContent?: string;
  changes: IChangeItem[];
  subSections?: ISectionComparison[];
}

/**
 * Side-by-side view data
 */
export interface ISideBySideView {
  leftVersion: IPolicyVersion;
  rightVersion: IPolicyVersion;
  alignedBlocks: IAlignedBlock[];
}

/**
 * Aligned content block for side-by-side view
 */
export interface IAlignedBlock {
  lineNumber: number;
  leftContent: string | null;
  rightContent: string | null;
  changeType: ChangeType;
  diffHtml?: {
    left: string;
    right: string;
  };
}

/**
 * Diff options
 */
export interface IDiffOptions {
  ignoreWhitespace?: boolean;
  ignoreCase?: boolean;
  contextLines?: number;
  wordLevel?: boolean;
  detectMoves?: boolean;
}

// ============================================================================
// SERVICE CLASS
// ============================================================================

/**
 * Policy Document Comparison Service
 */
export class PolicyDocumentComparisonService {
  private readonly sp: SPFI;
  private readonly siteUrl: string;

  // List names
  private readonly VERSIONS_LIST = ComparisonLists.VERSIONS;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly COMPARISON_HISTORY_LIST = ComparisonLists.COMPARISON_HISTORY;

  constructor(sp: SPFI, siteUrl: string) {
    this.sp = sp;
    this.siteUrl = siteUrl;
  }

  // ============================================================================
  // VERSION MANAGEMENT
  // ============================================================================

  /**
   * Get all versions of a policy
   */
  public async getPolicyVersions(policyId: number): Promise<IPolicyVersion[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.VERSIONS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .orderBy('VersionNumber', false)
        .top(100)();

      return items.map(item => this.mapToVersion(item));
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', `Failed to get versions for policy ${policyId}:`, error);
      return [];
    }
  }

  /**
   * Get specific version
   */
  public async getVersion(versionId: number): Promise<IPolicyVersion | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.VERSIONS_LIST)
        .items.getById(versionId)();

      return this.mapToVersion(item);
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', `Failed to get version ${versionId}:`, error);
      return null;
    }
  }

  /**
   * Create new version
   */
  public async createVersion(policyId: number, content: string, changeNotes?: string): Promise<number> {
    try {
      // Get current policy
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)() as IPolicy;

      // Get existing versions to determine version number
      const versions = await this.getPolicyVersions(policyId);
      const nextVersionNumber = versions.length > 0
        ? Math.max(...versions.map(v => v.versionNumber)) + 1
        : 1;

      const currentUser = await this.sp.web.currentUser();

      const result = await this.sp.web.lists
        .getByTitle(this.VERSIONS_LIST)
        .items.add({
          Title: `${policy.PolicyName} - v${nextVersionNumber}`,
          PolicyId: policyId,
          Version: `${nextVersionNumber}.0`,
          VersionNumber: nextVersionNumber,
          Content: content,
          Summary: policy.Description,
          EffectiveDate: policy.EffectiveDate,
          Status: policy.Status,
          ChangeNotes: changeNotes,
          WordCount: this.countWords(content),
          SectionCount: this.countSections(content),
          CreatedById: currentUser.Id,
          CreatedByName: currentUser.Title
        });

      logger.info('PolicyDocumentComparisonService', `Created version ${nextVersionNumber} for policy ${policyId}`);
      return result.data.Id;
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', `Failed to create version for policy ${policyId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // COMPARISON
  // ============================================================================

  /**
   * Compare two policy versions
   */
  public async compareVersions(
    sourceVersionId: number,
    targetVersionId: number,
    options?: IDiffOptions
  ): Promise<IComparisonResult> {
    try {
      const sourceVersion = await this.getVersion(sourceVersionId);
      const targetVersion = await this.getVersion(targetVersionId);

      if (!sourceVersion || !targetVersion) {
        throw new Error('One or both versions not found');
      }

      const changes = this.computeChanges(sourceVersion, targetVersion, options);
      const sectionComparisons = this.compareSections(sourceVersion.content, targetVersion.content);
      const summary = this.computeSummary(changes, sourceVersion, targetVersion);

      const currentUser = await this.sp.web.currentUser();

      const result: IComparisonResult = {
        sourceVersion,
        targetVersion,
        summary,
        changes,
        sectionComparisons,
        comparedDate: new Date().toISOString(),
        comparedById: currentUser.Id,
        comparedByName: currentUser.Title
      };

      // Save comparison to history
      await this.saveComparisonHistory(result);

      return result;
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', 'Failed to compare versions:', error);
      throw error;
    }
  }

  /**
   * Compare current policy with a specific version
   */
  public async compareWithVersion(policyId: number, versionId: number, options?: IDiffOptions): Promise<IComparisonResult> {
    try {
      // Get current policy content
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)() as IPolicy & { PolicyContent?: string };

      const currentUser = await this.sp.web.currentUser();

      // Create temporary "current" version object
      const currentVersion: IPolicyVersion = {
        id: 0,
        policyId: policyId,
        version: 'Current',
        versionNumber: 0,
        title: policy.PolicyName,
        content: policy.PolicyContent || '',
        summary: policy.Description,
        effectiveDate: policy.EffectiveDate instanceof Date
        ? policy.EffectiveDate.toISOString()
        : policy.EffectiveDate,
        createdDate: new Date().toISOString(),
        createdById: currentUser.Id,
        createdByName: currentUser.Title,
        status: policy.Status,
        wordCount: this.countWords(policy.PolicyContent || ''),
        sectionCount: this.countSections(policy.PolicyContent || '')
      };

      const previousVersion = await this.getVersion(versionId);
      if (!previousVersion) {
        throw new Error('Version not found');
      }

      const changes = this.computeChanges(previousVersion, currentVersion, options);
      const sectionComparisons = this.compareSections(previousVersion.content, currentVersion.content);
      const summary = this.computeSummary(changes, previousVersion, currentVersion);

      return {
        sourceVersion: previousVersion,
        targetVersion: currentVersion,
        summary,
        changes,
        sectionComparisons,
        comparedDate: new Date().toISOString(),
        comparedById: currentUser.Id,
        comparedByName: currentUser.Title
      };
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', 'Failed to compare with version:', error);
      throw error;
    }
  }

  /**
   * Get side-by-side view data
   */
  public async getSideBySideView(
    leftVersionId: number,
    rightVersionId: number,
    options?: IDiffOptions
  ): Promise<ISideBySideView> {
    try {
      const leftVersion = await this.getVersion(leftVersionId);
      const rightVersion = await this.getVersion(rightVersionId);

      if (!leftVersion || !rightVersion) {
        throw new Error('One or both versions not found');
      }

      const alignedBlocks = this.alignBlocks(
        leftVersion.content,
        rightVersion.content,
        options
      );

      return {
        leftVersion,
        rightVersion,
        alignedBlocks
      };
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', 'Failed to get side-by-side view:', error);
      throw error;
    }
  }

  // ============================================================================
  // DIFF ALGORITHMS
  // ============================================================================

  /**
   * Compute changes between two versions
   */
  private computeChanges(
    source: IPolicyVersion,
    target: IPolicyVersion,
    options?: IDiffOptions
  ): IChangeItem[] {
    const changes: IChangeItem[] = [];

    const sourceLines = this.splitIntoBlocks(source.content, options);
    const targetLines = this.splitIntoBlocks(target.content, options);

    // Use LCS-based diff algorithm
    const lcs = this.computeLCS(sourceLines, targetLines, options);

    let sourceIndex = 0;
    let targetIndex = 0;
    let changeId = 0;

    for (const match of lcs) {
      // Handle deletions (lines in source before the match)
      while (sourceIndex < match.sourceIndex) {
        changes.push({
          id: `change-${++changeId}`,
          type: ChangeType.Removed,
          blockType: this.detectBlockType(sourceLines[sourceIndex]),
          originalText: sourceLines[sourceIndex],
          originalPosition: sourceIndex,
          significance: this.assessSignificance(sourceLines[sourceIndex], ''),
          sectionNumber: this.extractSectionNumber(sourceLines[sourceIndex])
        });
        sourceIndex++;
      }

      // Handle additions (lines in target before the match)
      while (targetIndex < match.targetIndex) {
        changes.push({
          id: `change-${++changeId}`,
          type: ChangeType.Added,
          blockType: this.detectBlockType(targetLines[targetIndex]),
          newText: targetLines[targetIndex],
          newPosition: targetIndex,
          significance: this.assessSignificance('', targetLines[targetIndex]),
          sectionNumber: this.extractSectionNumber(targetLines[targetIndex])
        });
        targetIndex++;
      }

      // The lines match, skip them
      sourceIndex++;
      targetIndex++;
    }

    // Handle remaining deletions
    while (sourceIndex < sourceLines.length) {
      changes.push({
        id: `change-${++changeId}`,
        type: ChangeType.Removed,
        blockType: this.detectBlockType(sourceLines[sourceIndex]),
        originalText: sourceLines[sourceIndex],
        originalPosition: sourceIndex,
        significance: this.assessSignificance(sourceLines[sourceIndex], ''),
        sectionNumber: this.extractSectionNumber(sourceLines[sourceIndex])
      });
      sourceIndex++;
    }

    // Handle remaining additions
    while (targetIndex < targetLines.length) {
      changes.push({
        id: `change-${++changeId}`,
        type: ChangeType.Added,
        blockType: this.detectBlockType(targetLines[targetIndex]),
        newText: targetLines[targetIndex],
        newPosition: targetIndex,
        significance: this.assessSignificance('', targetLines[targetIndex]),
        sectionNumber: this.extractSectionNumber(targetLines[targetIndex])
      });
      targetIndex++;
    }

    // Detect modifications (adjacent add/remove pairs that are similar)
    if (options?.wordLevel) {
      return this.detectModifications(changes, options);
    }

    return changes;
  }

  /**
   * Compute Longest Common Subsequence
   */
  private computeLCS(
    source: string[],
    target: string[],
    options?: IDiffOptions
  ): { sourceIndex: number; targetIndex: number }[] {
    const m = source.length;
    const n = target.length;

    // Build LCS table
    const dp: number[][] = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        if (this.linesMatch(source[i - 1], target[j - 1], options)) {
          dp[i][j] = dp[i - 1][j - 1] + 1;
        } else {
          dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
        }
      }
    }

    // Backtrack to find matches
    const matches: { sourceIndex: number; targetIndex: number }[] = [];
    let i = m;
    let j = n;

    while (i > 0 && j > 0) {
      if (this.linesMatch(source[i - 1], target[j - 1], options)) {
        matches.unshift({ sourceIndex: i - 1, targetIndex: j - 1 });
        i--;
        j--;
      } else if (dp[i - 1][j] > dp[i][j - 1]) {
        i--;
      } else {
        j--;
      }
    }

    return matches;
  }

  /**
   * Check if two lines match based on options
   */
  private linesMatch(line1: string, line2: string, options?: IDiffOptions): boolean {
    let a = line1;
    let b = line2;

    if (options?.ignoreWhitespace) {
      a = a.replace(/\s+/g, ' ').trim();
      b = b.replace(/\s+/g, ' ').trim();
    }

    if (options?.ignoreCase) {
      a = a.toLowerCase();
      b = b.toLowerCase();
    }

    return a === b;
  }

  /**
   * Detect modifications by finding similar add/remove pairs
   */
  private detectModifications(changes: IChangeItem[], options?: IDiffOptions): IChangeItem[] {
    const result: IChangeItem[] = [];
    const usedIndices = new Set<number>();

    for (let i = 0; i < changes.length; i++) {
      if (usedIndices.has(i)) continue;

      const change = changes[i];

      if (change.type === ChangeType.Removed) {
        // Look for a nearby addition that's similar
        for (let j = i + 1; j < Math.min(i + 5, changes.length); j++) {
          if (usedIndices.has(j)) continue;

          const potentialMatch = changes[j];
          if (potentialMatch.type === ChangeType.Added) {
            const similarity = this.calculateSimilarity(
              change.originalText || '',
              potentialMatch.newText || ''
            );

            if (similarity > 0.5) {
              // This is a modification
              const wordChanges = this.computeWordDiff(
                change.originalText || '',
                potentialMatch.newText || '',
                options
              );

              result.push({
                ...change,
                id: change.id,
                type: ChangeType.Modified,
                newText: potentialMatch.newText,
                newPosition: potentialMatch.newPosition,
                wordChanges,
                significance: this.assessSignificance(change.originalText || '', potentialMatch.newText || '')
              });

              usedIndices.add(i);
              usedIndices.add(j);
              break;
            }
          }
        }

        if (!usedIndices.has(i)) {
          result.push(change);
          usedIndices.add(i);
        }
      } else {
        result.push(change);
        usedIndices.add(i);
      }
    }

    return result;
  }

  /**
   * Calculate similarity between two strings (Jaccard similarity)
   */
  private calculateSimilarity(str1: string, str2: string): number {
    const words1 = new Set(str1.toLowerCase().split(/\s+/));
    const words2 = new Set(str2.toLowerCase().split(/\s+/));

    const intersection = new Set(Array.from(words1).filter(x => words2.has(x)));
    const union = new Set([...Array.from(words1), ...Array.from(words2)]);

    return intersection.size / union.size;
  }

  /**
   * Compute word-level diff
   */
  private computeWordDiff(source: string, target: string, options?: IDiffOptions): IWordChange[] {
    const sourceWords = source.split(/\s+/);
    const targetWords = target.split(/\s+/);

    const changes: IWordChange[] = [];
    const lcs = this.computeLCS(sourceWords, targetWords, options);

    let sourceIdx = 0;
    let targetIdx = 0;
    let position = 0;

    for (const match of lcs) {
      while (sourceIdx < match.sourceIndex) {
        changes.push({
          type: ChangeType.Removed,
          text: sourceWords[sourceIdx],
          position: position++
        });
        sourceIdx++;
      }

      while (targetIdx < match.targetIndex) {
        changes.push({
          type: ChangeType.Added,
          text: targetWords[targetIdx],
          position: position++
        });
        targetIdx++;
      }

      changes.push({
        type: ChangeType.Unchanged,
        text: targetWords[targetIdx],
        position: position++
      });

      sourceIdx++;
      targetIdx++;
    }

    // Remaining removals
    while (sourceIdx < sourceWords.length) {
      changes.push({
        type: ChangeType.Removed,
        text: sourceWords[sourceIdx],
        position: position++
      });
      sourceIdx++;
    }

    // Remaining additions
    while (targetIdx < targetWords.length) {
      changes.push({
        type: ChangeType.Added,
        text: targetWords[targetIdx],
        position: position++
      });
      targetIdx++;
    }

    return changes;
  }

  /**
   * Compare sections between two documents
   */
  private compareSections(sourceContent: string, targetContent: string): ISectionComparison[] {
    const sourceSections = this.parseSections(sourceContent);
    const targetSections = this.parseSections(targetContent);

    const comparisons: ISectionComparison[] = [];

    // Track processed target sections
    const processedTargets = new Set<string>();

    // Compare source sections with target
    for (const sourceSection of sourceSections) {
      const matchingTarget = targetSections.find(t =>
        t.number === sourceSection.number ||
        this.calculateSimilarity(t.title, sourceSection.title) > 0.8
      );

      if (matchingTarget) {
        processedTargets.add(matchingTarget.number);

        const sectionChanges = this.computeChanges(
          { content: sourceSection.content } as IPolicyVersion,
          { content: matchingTarget.content } as IPolicyVersion
        );

        const status = sectionChanges.length > 0 ? ChangeType.Modified : ChangeType.Unchanged;

        comparisons.push({
          sectionNumber: sourceSection.number,
          sectionTitle: sourceSection.title,
          status,
          originalContent: sourceSection.content,
          newContent: matchingTarget.content,
          changes: sectionChanges
        });
      } else {
        // Section was removed
        comparisons.push({
          sectionNumber: sourceSection.number,
          sectionTitle: sourceSection.title,
          status: ChangeType.Removed,
          originalContent: sourceSection.content,
          changes: []
        });
      }
    }

    // Add new sections
    for (const targetSection of targetSections) {
      if (!processedTargets.has(targetSection.number)) {
        comparisons.push({
          sectionNumber: targetSection.number,
          sectionTitle: targetSection.title,
          status: ChangeType.Added,
          newContent: targetSection.content,
          changes: []
        });
      }
    }

    return comparisons.sort((a, b) => a.sectionNumber.localeCompare(b.sectionNumber));
  }

  /**
   * Parse content into sections
   */
  private parseSections(content: string): { number: string; title: string; content: string }[] {
    const sections: { number: string; title: string; content: string }[] = [];

    // Match section headings like "1. Introduction", "2.1 Scope", "Section 3: Purpose"
    const sectionRegex = /^((?:\d+\.)+\d*|Section\s+\d+:?)\s*(.+?)(?:\n|$)/gm;

    let match;
    let lastIndex = 0;
    let lastSection: { number: string; title: string; content: string } | null = null;

    while ((match = sectionRegex.exec(content)) !== null) {
      if (lastSection) {
        lastSection.content = content.substring(lastIndex, match.index).trim();
        sections.push(lastSection);
      }

      lastSection = {
        number: match[1].replace('Section ', '').replace(':', ''),
        title: match[2].trim(),
        content: ''
      };

      lastIndex = match.index + match[0].length;
    }

    // Add last section
    if (lastSection) {
      lastSection.content = content.substring(lastIndex).trim();
      sections.push(lastSection);
    }

    return sections;
  }

  /**
   * Align blocks for side-by-side view
   */
  private alignBlocks(leftContent: string, rightContent: string, options?: IDiffOptions): IAlignedBlock[] {
    const leftLines = leftContent.split('\n');
    const rightLines = rightContent.split('\n');

    const lcs = this.computeLCS(leftLines, rightLines, options);
    const alignedBlocks: IAlignedBlock[] = [];

    let leftIdx = 0;
    let rightIdx = 0;
    let lineNumber = 1;

    for (const match of lcs) {
      // Align removed lines (left only)
      while (leftIdx < match.sourceIndex) {
        alignedBlocks.push({
          lineNumber: lineNumber++,
          leftContent: leftLines[leftIdx],
          rightContent: null,
          changeType: ChangeType.Removed,
          diffHtml: {
            left: this.highlightRemoved(leftLines[leftIdx]),
            right: ''
          }
        });
        leftIdx++;
      }

      // Align added lines (right only)
      while (rightIdx < match.targetIndex) {
        alignedBlocks.push({
          lineNumber: lineNumber++,
          leftContent: null,
          rightContent: rightLines[rightIdx],
          changeType: ChangeType.Added,
          diffHtml: {
            left: '',
            right: this.highlightAdded(rightLines[rightIdx])
          }
        });
        rightIdx++;
      }

      // Aligned matching line
      alignedBlocks.push({
        lineNumber: lineNumber++,
        leftContent: leftLines[leftIdx],
        rightContent: rightLines[rightIdx],
        changeType: ChangeType.Unchanged
      });

      leftIdx++;
      rightIdx++;
    }

    // Remaining left lines
    while (leftIdx < leftLines.length) {
      alignedBlocks.push({
        lineNumber: lineNumber++,
        leftContent: leftLines[leftIdx],
        rightContent: null,
        changeType: ChangeType.Removed,
        diffHtml: {
          left: this.highlightRemoved(leftLines[leftIdx]),
          right: ''
        }
      });
      leftIdx++;
    }

    // Remaining right lines
    while (rightIdx < rightLines.length) {
      alignedBlocks.push({
        lineNumber: lineNumber++,
        leftContent: null,
        rightContent: rightLines[rightIdx],
        changeType: ChangeType.Added,
        diffHtml: {
          left: '',
          right: this.highlightAdded(rightLines[rightIdx])
        }
      });
      rightIdx++;
    }

    return alignedBlocks;
  }

  // ============================================================================
  // HTML GENERATION
  // ============================================================================

  /**
   * Generate unified diff HTML
   */
  public generateUnifiedDiffHtml(comparison: IComparisonResult): string {
    let html = `
      <div class="diff-unified">
        <div class="diff-header">
          <span class="version-old">- ${comparison.sourceVersion.version}</span>
          <span class="version-new">+ ${comparison.targetVersion.version}</span>
        </div>
        <div class="diff-content">
    `;

    for (const change of comparison.changes) {
      switch (change.type) {
        case ChangeType.Removed:
          html += `<div class="diff-line diff-removed">- ${this.escapeHtml(change.originalText || '')}</div>`;
          break;
        case ChangeType.Added:
          html += `<div class="diff-line diff-added">+ ${this.escapeHtml(change.newText || '')}</div>`;
          break;
        case ChangeType.Modified:
          html += `<div class="diff-line diff-removed">- ${this.escapeHtml(change.originalText || '')}</div>`;
          html += `<div class="diff-line diff-added">+ ${this.escapeHtml(change.newText || '')}</div>`;
          break;
      }
    }

    html += '</div></div>';
    return html;
  }

  /**
   * Generate side-by-side diff HTML
   */
  public generateSideBySideHtml(sideBySide: ISideBySideView): string {
    let html = `
      <div class="diff-sidebyside">
        <div class="diff-header">
          <div class="diff-col-header">${sideBySide.leftVersion.version}</div>
          <div class="diff-col-header">${sideBySide.rightVersion.version}</div>
        </div>
        <div class="diff-content">
    `;

    for (const block of sideBySide.alignedBlocks) {
      const leftClass = block.changeType === ChangeType.Removed ? 'diff-removed' : '';
      const rightClass = block.changeType === ChangeType.Added ? 'diff-added' : '';

      html += `
        <div class="diff-row">
          <div class="diff-line-number">${block.lineNumber}</div>
          <div class="diff-left ${leftClass}">${block.diffHtml?.left || this.escapeHtml(block.leftContent || '')}</div>
          <div class="diff-right ${rightClass}">${block.diffHtml?.right || this.escapeHtml(block.rightContent || '')}</div>
        </div>
      `;
    }

    html += '</div></div>';
    return html;
  }

  /**
   * Generate inline diff with word highlighting
   */
  public generateInlineDiffHtml(wordChanges: IWordChange[]): string {
    let html = '';

    for (const change of wordChanges) {
      switch (change.type) {
        case ChangeType.Removed:
          html += `<del class="diff-word-removed">${this.escapeHtml(change.text)}</del> `;
          break;
        case ChangeType.Added:
          html += `<ins class="diff-word-added">${this.escapeHtml(change.text)}</ins> `;
          break;
        case ChangeType.Unchanged:
          html += `${this.escapeHtml(change.text)} `;
          break;
      }
    }

    return html.trim();
  }

  // ============================================================================
  // HISTORY & ANALYTICS
  // ============================================================================

  /**
   * Save comparison to history
   */
  private async saveComparisonHistory(comparison: IComparisonResult): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.COMPARISON_HISTORY_LIST)
        .items.add({
          Title: `${comparison.sourceVersion.title} v${comparison.sourceVersion.version} vs v${comparison.targetVersion.version}`,
          PolicyId: comparison.sourceVersion.policyId,
          SourceVersionId: comparison.sourceVersion.id,
          TargetVersionId: comparison.targetVersion.id,
          TotalChanges: comparison.summary.totalChanges,
          Additions: comparison.summary.additions,
          Deletions: comparison.summary.deletions,
          Modifications: comparison.summary.modifications,
          PercentageChanged: comparison.summary.percentageChanged,
          ComparedById: comparison.comparedById,
          ComparedByName: comparison.comparedByName,
          ComparedDate: comparison.comparedDate
        });
    } catch (error) {
      logger.warn('PolicyDocumentComparisonService', 'Failed to save comparison history:', error);
    }
  }

  /**
   * Get comparison history for a policy
   */
  public async getComparisonHistory(policyId: number): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.COMPARISON_HISTORY_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .orderBy('ComparedDate', false)
        .top(50)();

      return items;
    } catch (error) {
      logger.error('PolicyDocumentComparisonService', 'Failed to get comparison history:', error);
      return [];
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Split content into comparable blocks
   */
  private splitIntoBlocks(content: string, options?: IDiffOptions): string[] {
    let lines = content.split('\n');

    if (options?.ignoreWhitespace) {
      lines = lines.map(line => line.replace(/\s+/g, ' ').trim());
    }

    return lines.filter(line => line.length > 0);
  }

  /**
   * Detect block type from content
   */
  private detectBlockType(content: string): BlockType {
    if (/^#{1,6}\s/.test(content) || /^[\d.]+\s/.test(content)) {
      return BlockType.Heading;
    }
    if (/^[-*•]\s/.test(content) || /^\d+\.\s/.test(content)) {
      return BlockType.List;
    }
    if (/\|.*\|/.test(content)) {
      return BlockType.Table;
    }
    if (/!\[.*\]\(.*\)/.test(content)) {
      return BlockType.Image;
    }
    return BlockType.Paragraph;
  }

  /**
   * Extract section number from content
   */
  private extractSectionNumber(content: string): string | undefined {
    const match = content.match(/^([\d.]+)/);
    return match ? match[1] : undefined;
  }

  /**
   * Assess change significance
   */
  private assessSignificance(originalText: string, newText: string): 'Major' | 'Minor' | 'Cosmetic' {
    // Cosmetic: only whitespace/punctuation changes
    const normalizedOriginal = originalText.replace(/[\s.,;:!?'"()[\]{}]/g, '').toLowerCase();
    const normalizedNew = newText.replace(/[\s.,;:!?'"()[\]{}]/g, '').toLowerCase();

    if (normalizedOriginal === normalizedNew) {
      return 'Cosmetic';
    }

    // Major: significant content change (>30% words changed)
    const originalWords = originalText.toLowerCase().split(/\s+/).filter(w => w.length > 0);
    const newWords = newText.toLowerCase().split(/\s+/).filter(w => w.length > 0);

    const commonWords = originalWords.filter(w => newWords.includes(w));
    const totalWords = Math.max(originalWords.length, newWords.length);
    const changeRatio = 1 - (commonWords.length / totalWords);

    if (changeRatio > 0.3 || Math.abs(originalWords.length - newWords.length) > 10) {
      return 'Major';
    }

    return 'Minor';
  }

  /**
   * Compute summary statistics
   */
  private computeSummary(
    changes: IChangeItem[],
    source: IPolicyVersion,
    target: IPolicyVersion
  ): IComparisonResult['summary'] {
    const additions = changes.filter(c => c.type === ChangeType.Added).length;
    const deletions = changes.filter(c => c.type === ChangeType.Removed).length;
    const modifications = changes.filter(c => c.type === ChangeType.Modified).length;

    const majorChanges = changes.filter(c => c.significance === 'Major').length;
    const minorChanges = changes.filter(c => c.significance === 'Minor').length;
    const cosmeticChanges = changes.filter(c => c.significance === 'Cosmetic').length;

    const sourceWords = source.wordCount || this.countWords(source.content);
    const targetWords = target.wordCount || this.countWords(target.content);
    const wordCountChange = targetWords - sourceWords;

    // Calculate percentage changed
    const totalSourceLines = source.content.split('\n').length;
    const percentageChanged = totalSourceLines > 0
      ? Math.round((changes.length / totalSourceLines) * 100)
      : 0;

    return {
      totalChanges: changes.length,
      additions,
      deletions,
      modifications,
      majorChanges,
      minorChanges,
      cosmeticChanges,
      wordCountChange,
      percentageChanged
    };
  }

  /**
   * Count words in content
   */
  private countWords(content: string): number {
    return content.split(/\s+/).filter(w => w.length > 0).length;
  }

  /**
   * Count sections in content
   */
  private countSections(content: string): number {
    const sectionRegex = /^((?:\d+\.)+\d*|Section\s+\d+:?)\s/gm;
    const matches = content.match(sectionRegex);
    return matches ? matches.length : 0;
  }

  /**
   * Escape HTML special characters
   */
  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /**
   * Highlight removed text
   */
  private highlightRemoved(text: string): string {
    return `<span class="diff-highlight-removed">${this.escapeHtml(text)}</span>`;
  }

  /**
   * Highlight added text
   */
  private highlightAdded(text: string): string {
    return `<span class="diff-highlight-added">${this.escapeHtml(text)}</span>`;
  }

  /**
   * Map SharePoint item to Version
   */
  private mapToVersion(item: any): IPolicyVersion {
    return {
      id: item.Id,
      policyId: item.PolicyId,
      version: item.Version,
      versionNumber: item.VersionNumber,
      title: item.Title,
      content: item.Content || '',
      contentHtml: item.ContentHtml,
      summary: item.Summary,
      effectiveDate: item.EffectiveDate,
      createdDate: item.Created,
      createdById: item.CreatedById,
      createdByName: item.CreatedByName,
      changeNotes: item.ChangeNotes,
      status: item.Status,
      wordCount: item.WordCount,
      sectionCount: item.SectionCount
    };
  }
}

// Export factory
export const createPolicyDocumentComparisonService = (sp: SPFI, siteUrl: string): PolicyDocumentComparisonService => {
  return new PolicyDocumentComparisonService(sp, siteUrl);
};
