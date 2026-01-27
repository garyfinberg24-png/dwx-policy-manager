// @ts-nocheck
/**
 * PowerPoint Generation Service
 * Creates PowerPoint presentations with slides and content
 */

import PptxGenJS from 'pptxgenjs';
import { ITemplateDataContext } from './DocxTemplateProcessor';
import { logger } from './LoggingService';

/**
 * PowerPoint generation result
 */
export interface IPowerPointGenerationResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  error?: string;
}

/**
 * PowerPoint document options
 */
export interface IPowerPointDocumentOptions {
  /** Presentation title */
  title?: string;
  /** Author name */
  author?: string;
  /** Company name */
  company?: string;
  /** Subject */
  subject?: string;
  /** Layout: 16x9 or 4x3 */
  layout?: '16x9' | '4x3';
  /** Primary brand color (hex) */
  primaryColor?: string;
  /** Secondary brand color (hex) */
  secondaryColor?: string;
}

/**
 * Slide content types
 */
export type SlideType = 'title' | 'content' | 'section' | 'twoColumn' | 'table' | 'chart' | 'blank';

/**
 * Slide configuration
 */
export interface ISlideConfig {
  /** Slide type */
  type: SlideType;
  /** Slide title */
  title?: string;
  /** Slide subtitle */
  subtitle?: string;
  /** Body content (for content slides) */
  body?: string | string[];
  /** Left column content (for twoColumn) */
  leftColumn?: string | string[];
  /** Right column content (for twoColumn) */
  rightColumn?: string[];
  /** Table data (for table slides) */
  tableData?: {
    headers: string[];
    rows: string[][];
  };
  /** Notes for the slide */
  notes?: string;
}

/**
 * PowerPoint Generation Service
 */
export class PowerPointGenerationService {
  private primaryColor: string = '0078D4';
  private secondaryColor: string = '323130';

  /**
   * Generate a PowerPoint presentation
   */
  public async generatePresentation(
    slides: ISlideConfig[],
    options?: IPowerPointDocumentOptions
  ): Promise<IPowerPointGenerationResult> {
    try {
      const pptx = new PptxGenJS();

      // Set presentation properties
      pptx.author = options?.author || 'JML Document Builder';
      pptx.company = options?.company || '';
      pptx.title = options?.title || 'Presentation';
      pptx.subject = options?.subject || '';

      // Set layout
      if (options?.layout === '4x3') {
        pptx.defineLayout({ name: 'LAYOUT_4x3', width: 10, height: 7.5 });
        pptx.layout = 'LAYOUT_4x3';
      } else {
        pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
        pptx.layout = 'LAYOUT_16x9';
      }

      // Set colors
      if (options?.primaryColor) {
        this.primaryColor = options.primaryColor.replace('#', '');
      }
      if (options?.secondaryColor) {
        this.secondaryColor = options.secondaryColor.replace('#', '');
      }

      // Add slides
      for (let i = 0; i < slides.length; i++) {
        this.addSlide(pptx, slides[i]);
      }

      // Generate presentation
      const data = await pptx.write({ outputType: 'blob' });
      const blob = data as Blob;

      const fileName = this.generateFileName(options);

      logger.info('PowerPointGenerationService', `PowerPoint generated: ${fileName}`);

      return {
        success: true,
        blob,
        fileName
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('PowerPointGenerationService', 'Failed to generate PowerPoint:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Generate an employee onboarding presentation
   */
  public async generateOnboardingPresentation(
    dataContext: Partial<ITemplateDataContext>,
    additionalSlides?: ISlideConfig[],
    options?: IPowerPointDocumentOptions
  ): Promise<IPowerPointGenerationResult> {
    const slides: ISlideConfig[] = [
      {
        type: 'title',
        title: `Welcome ${dataContext.employee?.name || 'New Employee'}!`,
        subtitle: `${dataContext.employee?.jobTitle || 'Team Member'} | ${dataContext.employee?.department || 'Department'}`
      },
      {
        type: 'section',
        title: 'Getting Started',
        subtitle: 'Your first steps at the company'
      },
      {
        type: 'content',
        title: 'Your Information',
        body: [
          `Name: ${dataContext.employee?.name || 'N/A'}`,
          `Email: ${dataContext.employee?.email || 'N/A'}`,
          `Department: ${dataContext.employee?.department || 'N/A'}`,
          `Job Title: ${dataContext.employee?.jobTitle || 'N/A'}`,
          `Start Date: ${dataContext.employee?.startDate || 'N/A'}`,
          `Manager: ${dataContext.manager?.name || 'N/A'}`
        ]
      },
      {
        type: 'twoColumn',
        title: 'Key Contacts',
        leftColumn: [
          'Your Manager',
          `${dataContext.manager?.name || 'TBD'}`,
          `${dataContext.manager?.email || ''}`
        ],
        rightColumn: [
          'HR Contact',
          'HR Department',
          'hr@company.com'
        ]
      },
      {
        type: 'content',
        title: 'First Week Checklist',
        body: [
          '✓ Complete onboarding paperwork',
          '✓ Set up IT equipment and accounts',
          '✓ Meet your team members',
          '✓ Review company policies',
          '✓ Attend orientation sessions',
          '✓ Schedule 1:1 with your manager'
        ]
      }
    ];

    // Add any additional slides
    if (additionalSlides) {
      slides.push(...additionalSlides);
    }

    // Add closing slide
    slides.push({
      type: 'section',
      title: 'Welcome to the Team!',
      subtitle: 'We\'re excited to have you on board'
    });

    return this.generatePresentation(slides, {
      ...options,
      title: options?.title || `Onboarding - ${dataContext.employee?.name || 'New Employee'}`
    });
  }

  /**
   * Generate a process summary presentation
   */
  public async generateProcessSummaryPresentation(
    processes: Array<{
      type: string;
      employeeName: string;
      department: string;
      status: string;
      progress: number;
    }>,
    options?: IPowerPointDocumentOptions
  ): Promise<IPowerPointGenerationResult> {
    const slides: ISlideConfig[] = [
      {
        type: 'title',
        title: 'JML Process Summary',
        subtitle: `Generated on ${new Date().toLocaleDateString()}`
      },
      {
        type: 'content',
        title: 'Overview',
        body: [
          `Total Processes: ${processes.length}`,
          `Joiners: ${processes.filter(p => p.type === 'Joiner').length}`,
          `Movers: ${processes.filter(p => p.type === 'Mover').length}`,
          `Leavers: ${processes.filter(p => p.type === 'Leaver').length}`
        ]
      },
      {
        type: 'table',
        title: 'Active Processes',
        tableData: {
          headers: ['Employee', 'Type', 'Department', 'Status', 'Progress'],
          rows: processes.slice(0, 10).map(p => [
            p.employeeName,
            p.type,
            p.department,
            p.status,
            `${p.progress}%`
          ])
        }
      }
    ];

    return this.generatePresentation(slides, {
      ...options,
      title: options?.title || 'JML Process Summary'
    });
  }

  /**
   * Add a slide to the presentation
   */
  private addSlide(pptx: PptxGenJS, config: ISlideConfig): void {
    const slide = pptx.addSlide();

    // Add notes if provided
    if (config.notes) {
      slide.addNotes(config.notes);
    }

    switch (config.type) {
      case 'title':
        this.createTitleSlide(slide, config);
        break;
      case 'section':
        this.createSectionSlide(slide, config);
        break;
      case 'content':
        this.createContentSlide(slide, config);
        break;
      case 'twoColumn':
        this.createTwoColumnSlide(slide, config);
        break;
      case 'table':
        this.createTableSlide(slide, config);
        break;
      case 'blank':
        // Just add title if provided
        if (config.title) {
          slide.addText(config.title, {
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 0.5,
            fontSize: 24,
            bold: true,
            color: this.secondaryColor
          });
        }
        break;
    }
  }

  /**
   * Create title slide
   */
  private createTitleSlide(slide: PptxGenJS.Slide, config: ISlideConfig): void {
    // Background shape
    slide.addShape('rect', {
      x: 0,
      y: 0,
      w: '100%',
      h: '100%',
      fill: { color: this.primaryColor }
    });

    // Title
    if (config.title) {
      slide.addText(config.title, {
        x: 0.5,
        y: 2,
        w: 9,
        h: 1,
        fontSize: 44,
        bold: true,
        color: 'FFFFFF',
        align: 'center'
      });
    }

    // Subtitle
    if (config.subtitle) {
      slide.addText(config.subtitle, {
        x: 0.5,
        y: 3.2,
        w: 9,
        h: 0.5,
        fontSize: 20,
        color: 'FFFFFF',
        align: 'center'
      });
    }
  }

  /**
   * Create section slide
   */
  private createSectionSlide(slide: PptxGenJS.Slide, config: ISlideConfig): void {
    // Left accent bar
    slide.addShape('rect', {
      x: 0,
      y: 0,
      w: 0.2,
      h: '100%',
      fill: { color: this.primaryColor }
    });

    // Title
    if (config.title) {
      slide.addText(config.title, {
        x: 0.5,
        y: 2,
        w: 9,
        h: 1,
        fontSize: 36,
        bold: true,
        color: this.secondaryColor
      });
    }

    // Subtitle
    if (config.subtitle) {
      slide.addText(config.subtitle, {
        x: 0.5,
        y: 3,
        w: 9,
        h: 0.5,
        fontSize: 18,
        color: '605E5C'
      });
    }
  }

  /**
   * Create content slide
   */
  private createContentSlide(slide: PptxGenJS.Slide, config: ISlideConfig): void {
    // Title
    if (config.title) {
      slide.addText(config.title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.6,
        fontSize: 28,
        bold: true,
        color: this.secondaryColor
      });
    }

    // Body content
    if (config.body) {
      const bodyContent = Array.isArray(config.body) ? config.body : [config.body];
      const bulletPoints = bodyContent.map(item => ({
        text: item,
        options: { bullet: { type: 'bullet' as const }, indentLevel: 0 }
      }));

      slide.addText(bulletPoints, {
        x: 0.5,
        y: 1.1,
        w: 9,
        h: 4,
        fontSize: 18,
        color: this.secondaryColor,
        valign: 'top'
      });
    }
  }

  /**
   * Create two-column slide
   */
  private createTwoColumnSlide(slide: PptxGenJS.Slide, config: ISlideConfig): void {
    // Title
    if (config.title) {
      slide.addText(config.title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.6,
        fontSize: 28,
        bold: true,
        color: this.secondaryColor
      });
    }

    // Left column
    if (config.leftColumn) {
      const leftContent = Array.isArray(config.leftColumn) ? config.leftColumn : [config.leftColumn];
      slide.addText(leftContent.join('\n'), {
        x: 0.5,
        y: 1.1,
        w: 4.3,
        h: 4,
        fontSize: 16,
        color: this.secondaryColor,
        valign: 'top'
      });
    }

    // Right column
    if (config.rightColumn) {
      slide.addText(config.rightColumn.join('\n'), {
        x: 5.2,
        y: 1.1,
        w: 4.3,
        h: 4,
        fontSize: 16,
        color: this.secondaryColor,
        valign: 'top'
      });
    }

    // Divider line
    slide.addShape('line', {
      x: 4.9,
      y: 1.1,
      w: 0,
      h: 3.5,
      line: { color: 'E1DFDD', width: 1 }
    });
  }

  /**
   * Create table slide
   */
  private createTableSlide(slide: PptxGenJS.Slide, config: ISlideConfig): void {
    // Title
    if (config.title) {
      slide.addText(config.title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.6,
        fontSize: 28,
        bold: true,
        color: this.secondaryColor
      });
    }

    // Table
    if (config.tableData) {
      const tableRows: PptxGenJS.TableRow[] = [];

      // Header row
      tableRows.push(
        config.tableData.headers.map(header => ({
          text: header,
          options: {
            fill: { color: this.primaryColor },
            color: 'FFFFFF',
            bold: true,
            align: 'center' as const
          }
        }))
      );

      // Data rows
      for (let i = 0; i < config.tableData.rows.length; i++) {
        tableRows.push(
          config.tableData.rows[i].map(cell => ({
            text: cell,
            options: {
              fill: { color: i % 2 === 0 ? 'FFFFFF' : 'FAF9F8' },
              color: this.secondaryColor
            }
          }))
        );
      }

      slide.addTable(tableRows, {
        x: 0.5,
        y: 1.1,
        w: 9,
        colW: 9 / config.tableData.headers.length,
        fontSize: 12,
        border: { type: 'solid', pt: 0.5, color: 'E1DFDD' }
      });
    }
  }

  /**
   * Generate file name
   */
  private generateFileName(options?: IPowerPointDocumentOptions): string {
    const timestamp = new Date().toISOString().slice(0, 10);
    const baseName = options?.title?.replace(/\s+/g, '_') || 'Presentation';
    return `${baseName}_${timestamp}.pptx`;
  }

  /**
   * Download PowerPoint file in browser
   */
  public downloadPresentation(result: IPowerPointGenerationResult): void {
    if (!result.success || !result.blob) {
      logger.error('PowerPointGenerationService', 'Cannot download: PowerPoint generation failed');
      return;
    }

    const url = URL.createObjectURL(result.blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = result.fileName || 'presentation.pptx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    logger.info('PowerPointGenerationService', `PowerPoint downloaded: ${result.fileName}`);
  }
}

export const powerPointGenerationService = new PowerPointGenerationService();
