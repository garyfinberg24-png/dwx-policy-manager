// @ts-nocheck
// Document Template Service
// Manages document templates with placeholders and dynamic content

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import {
  IJmlDocumentTemplate,
  ITemplatePlaceholder,
  DocumentType,
  SignatureProvider,
  IJmlProcess
} from '../models';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';
import { DocxTemplateProcessor, IDocxProcessingResult, ITemplateDataContext } from './DocxTemplateProcessor';

export class DocumentTemplateService {
  private sp: SPFI;
  private readonly templateLibrary: string;
  private readonly docxProcessor: DocxTemplateProcessor;

  constructor(sp: SPFI, templateLibraryUrl?: string) {
    this.sp = sp;
    this.templateLibrary = templateLibraryUrl || 'JML_DocumentTemplates';
    this.docxProcessor = new DocxTemplateProcessor();
  }

  /**
   * Get all document templates
   */
  public async getTemplates(documentType?: DocumentType): Promise<IJmlDocumentTemplate[]> {
    try {
      // Validate enum value if provided
      if (documentType) {
        ValidationUtils.validateEnum(documentType, DocumentType, 'DocumentType');
      }

      // Build secure filter
      let filter = 'IsActive eq true';
      if (documentType) {
        const docTypeFilter = ValidationUtils.buildFilter('DocumentType', 'eq', documentType);
        filter = `${filter} and ${docTypeFilter}`;
      }

      const items = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.select(
          '*',
          'File/ServerRelativeUrl',
          'CreatedBy/Title',
          'CreatedBy/EMail',
          'ModifiedBy/Title',
          'ModifiedBy/EMail'
        )
        .expand('File', 'CreatedBy', 'ModifiedBy')
        .filter(filter)
        .orderBy('Title')();

      const templates: IJmlDocumentTemplate[] = [];

      for (let i = 0; i < items.length; i++) {
        templates.push(this.mapToTemplate(items[i]));
      }

      return templates;
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to get templates, returning sample templates:', error);
      // Return sample templates if the list doesn't exist
      return this.getSampleTemplates(documentType);
    }
  }

  /**
   * Get sample templates for testing/demo purposes
   */
  private getSampleTemplates(documentType?: DocumentType): IJmlDocumentTemplate[] {
    const allTemplates: IJmlDocumentTemplate[] = [
      {
        Id: 1,
        Title: 'Employment Offer Letter',
        Description: 'Formal offer letter with compensation and benefits details',
        DocumentType: DocumentType.OfferLetter,
        ProcessTypes: ['Joiner'],
        Placeholders: [
          { Key: 'CandidateName', Label: 'Candidate Name', DataType: 'text', Required: true, DefaultValue: '' },
          { Key: 'JobTitle', Label: 'Job Title', DataType: 'text', Required: true, DefaultValue: '' },
          { Key: 'Salary', Label: 'Annual Salary', DataType: 'number', Required: true, DefaultValue: '' },
          { Key: 'StartDate', Label: 'Start Date', DataType: 'date', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/offer-letter.docx',
        RequiresSignature: true,
        SignatureProvider: SignatureProvider.DocuSign,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      },
      {
        Id: 2,
        Title: 'IT Equipment Assignment Form',
        Description: 'Document for tracking assigned IT assets and equipment',
        DocumentType: DocumentType.EquipmentForm,
        ProcessTypes: ['Joiner', 'Mover'],
        Placeholders: [
          { Key: 'EmployeeName', Label: 'Employee Name', DataType: 'user', Required: true, DefaultValue: '' },
          { Key: 'EmployeeID', Label: 'Employee ID', DataType: 'text', Required: true, DefaultValue: '' },
          { Key: 'Equipment', Label: 'Equipment List', DataType: 'text', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/equipment-form.docx',
        RequiresSignature: true,
        SignatureProvider: SignatureProvider.DocuSign,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      },
      {
        Id: 3,
        Title: 'Confidentiality Agreement (NDA)',
        Description: 'Non-disclosure agreement for protecting company information',
        DocumentType: DocumentType.NDAAgreement,
        ProcessTypes: ['Joiner'],
        Placeholders: [
          { Key: 'EmployeeName', Label: 'Employee Name', DataType: 'user', Required: true, DefaultValue: '' },
          { Key: 'EffectiveDate', Label: 'Effective Date', DataType: 'date', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/nda.docx',
        RequiresSignature: true,
        SignatureProvider: SignatureProvider.AdobeSign,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      },
      {
        Id: 4,
        Title: 'Exit Interview Form',
        Description: 'Structured exit interview questionnaire for departing employees',
        DocumentType: DocumentType.ExitForm,
        ProcessTypes: ['Leaver'],
        Placeholders: [
          { Key: 'EmployeeName', Label: 'Employee Name', DataType: 'user', Required: true, DefaultValue: '' },
          { Key: 'LastWorkingDay', Label: 'Last Working Day', DataType: 'date', Required: true, DefaultValue: '' },
          { Key: 'Department', Label: 'Department', DataType: 'department', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/exit-interview.docx',
        RequiresSignature: false,
        SignatureProvider: undefined,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      },
      {
        Id: 5,
        Title: 'Company Policy Acknowledgement',
        Description: 'Acknowledgement form for company policies and code of conduct',
        DocumentType: DocumentType.PolicyDocument,
        ProcessTypes: ['Joiner', 'Mover'],
        Placeholders: [
          { Key: 'EmployeeName', Label: 'Employee Name', DataType: 'user', Required: true, DefaultValue: '' },
          { Key: 'EmployeeID', Label: 'Employee ID', DataType: 'text', Required: true, DefaultValue: '' },
          { Key: 'PolicyDate', Label: 'Policy Date', DataType: 'date', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/policy-acknowledgement.docx',
        RequiresSignature: true,
        SignatureProvider: SignatureProvider.AdobeSign,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      },
      {
        Id: 6,
        Title: 'Employee Handbook Acknowledgment',
        Description: 'Acknowledgment that employee has received and read the handbook',
        DocumentType: DocumentType.HandbookAcknowledgment,
        ProcessTypes: ['Joiner'],
        Placeholders: [
          { Key: 'EmployeeName', Label: 'Employee Name', DataType: 'user', Required: true, DefaultValue: '' },
          { Key: 'SignDate', Label: 'Signature Date', DataType: 'date', Required: true, DefaultValue: '' }
        ],
        TemplateUrl: '/templates/handbook-ack.docx',
        RequiresSignature: true,
        SignatureProvider: SignatureProvider.Internal,
        IsActive: true,
        Created: new Date(),
        CreatedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        },
        Modified: new Date(),
        ModifiedBy: {
          Id: 1,
          Title: 'System',
          EMail: 'system@example.com'
        }
      }
    ];

    // Filter by document type if specified
    if (documentType) {
      return allTemplates.filter(t => t.DocumentType === documentType);
    }

    return allTemplates;
  }

  /**
   * Get template by ID
   */
  public async getTemplateById(templateId: number): Promise<IJmlDocumentTemplate> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select(
          '*',
          'File/ServerRelativeUrl',
          'CreatedBy/Title',
          'CreatedBy/EMail',
          'ModifiedBy/Title',
          'ModifiedBy/EMail'
        )
        .expand('File', 'CreatedBy', 'ModifiedBy')();

      return this.mapToTemplate(item);
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to get template:', error);
      throw error;
    }
  }

  /**
   * Create document from template with placeholder replacement
   * Uses docxtemplater for proper DOCX processing
   */
  public async createFromTemplate(
    templateId: number,
    process: IJmlProcess,
    placeholderValues: { [key: string]: string },
    companyInfo?: { name: string; address: string; phone: string }
  ): Promise<IDocxProcessingResult> {
    try {
      const template = await this.getTemplateById(templateId);

      // Build the template URL
      const templateUrl = template.TemplateUrl.startsWith('http')
        ? template.TemplateUrl
        : `${window.location.origin}${template.TemplateUrl}`;

      // Build data context from process and custom values
      const dataContext = this.docxProcessor.buildDataContext(
        process,
        placeholderValues,
        companyInfo
      );

      // Validate the template before processing
      const response = await fetch(templateUrl);
      if (!response.ok) {
        throw new Error(`Failed to fetch template: ${response.statusText}`);
      }
      const templateBlob = await response.blob();

      const validation = await this.docxProcessor.validateTemplate(
        templateBlob,
        template.Placeholders,
        dataContext
      );

      if (!validation.isValid) {
        logger.warn('DocumentTemplateService', 'Template validation warnings:', { missingFields: validation.missingFields });
      }

      // Process the template with docxtemplater
      const result = await this.docxProcessor.processTemplate(
        templateBlob,
        dataContext,
        this.generateDocumentFileName(template, process)
      );

      if (!result.success) {
        throw new Error(result.error || 'Failed to process template');
      }

      logger.info('DocumentTemplateService', `Document generated: ${result.fileName}`);
      return result;
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to create document from template:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create document'
      };
    }
  }

  /**
   * Create document from template (legacy method for backward compatibility)
   * @deprecated Use createFromTemplate instead
   */
  public async createFromTemplateLegacy(
    templateId: number,
    process: IJmlProcess,
    placeholderValues: { [key: string]: string }
  ): Promise<Blob> {
    try {
      const template = await this.getTemplateById(templateId);

      // Download template file
      const templateUrl = `${window.location.origin}${template.TemplateUrl}`;
      const response = await fetch(templateUrl);
      const templateBlob = await response.blob();

      // Read template content
      const content = await templateBlob.text();

      // Replace placeholders
      const processedContent = this.replacePlaceholders(
        content,
        template.Placeholders,
        process,
        placeholderValues
      );

      // Create new blob with processed content
      return new Blob([processedContent], { type: templateBlob.type });
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to create document from template:', error);
      throw error;
    }
  }

  /**
   * Generate a document file name based on template and process
   */
  private generateDocumentFileName(template: IJmlDocumentTemplate, process: IJmlProcess): string {
    const timestamp = new Date().toISOString().slice(0, 10);
    const employeeName = process.EmployeeName.replace(/\s+/g, '_');
    const templateName = template.Title.replace(/\s+/g, '_');
    return `${templateName}_${employeeName}_${timestamp}.docx`;
  }

  /**
   * Extract placeholders from a template file
   */
  public async extractTemplatePlaceholders(templateUrl: string): Promise<string[]> {
    try {
      const fullUrl = templateUrl.startsWith('http')
        ? templateUrl
        : `${window.location.origin}${templateUrl}`;

      const response = await fetch(fullUrl);
      if (!response.ok) {
        throw new Error(`Failed to fetch template: ${response.statusText}`);
      }

      const templateBlob = await response.blob();
      return this.docxProcessor.extractPlaceholders(templateBlob);
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to extract placeholders:', error);
      return [];
    }
  }

  /**
   * Validate template data before generation
   */
  public async validateTemplateData(
    templateId: number,
    process: IJmlProcess,
    placeholderValues: { [key: string]: string }
  ): Promise<{ isValid: boolean; missingFields: string[]; warnings: string[] }> {
    try {
      const template = await this.getTemplateById(templateId);

      const templateUrl = template.TemplateUrl.startsWith('http')
        ? template.TemplateUrl
        : `${window.location.origin}${template.TemplateUrl}`;

      const response = await fetch(templateUrl);
      if (!response.ok) {
        throw new Error(`Failed to fetch template: ${response.statusText}`);
      }

      const templateBlob = await response.blob();
      const dataContext = this.docxProcessor.buildDataContext(process, placeholderValues);

      return this.docxProcessor.validateTemplate(
        templateBlob,
        template.Placeholders,
        dataContext
      );
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to validate template data:', error);
      return {
        isValid: false,
        missingFields: ['Validation failed'],
        warnings: []
      };
    }
  }

  /**
   * Replace placeholders in template content
   */
  private replacePlaceholders(
    content: string,
    placeholders: ITemplatePlaceholder[],
    process: IJmlProcess,
    customValues: { [key: string]: string }
  ): string {
    let result = content;

    // Standard process placeholders
    const standardPlaceholders: { [key: string]: string } = {
      '{EmployeeName}': process.EmployeeName,
      '{EmployeeEmail}': process.EmployeeEmail || '',
      '{Department}': process.Department,
      '{StartDate}': process.StartDate ? this.formatDate(process.StartDate) : '',
      '{ManagerName}': process.Manager?.Title || '',
      '{ManagerEmail}': process.Manager?.EMail || '',
      '{ProcessType}': process.ProcessType,
      '{Today}': this.formatDate(new Date())
    };

    // Apply standard placeholders
    const standardKeys = Object.keys(standardPlaceholders);
    for (let i = 0; i < standardKeys.length; i++) {
      const key = standardKeys[i];
      const value = standardPlaceholders[key];
      result = result.split(key).join(value);
    }

    // Apply custom placeholders
    for (let i = 0; i < placeholders.length; i++) {
      const placeholder = placeholders[i];
      const key = `{${placeholder.Key}}`;
      const value = customValues[placeholder.Key] || placeholder.DefaultValue || '';
      result = result.split(key).join(value);
    }

    return result;
  }

  /**
   * Create template
   */
  public async createTemplate(
    file: File,
    title: string,
    documentType: DocumentType,
    processTypes: string[],
    placeholders: ITemplatePlaceholder[],
    requiresSignature: boolean
  ): Promise<IJmlDocumentTemplate> {
    try {
      // Upload template file
      const uploadResult = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .rootFolder
        .files.addUsingPath(file.name, file, { Overwrite: true });

      // Create metadata
      const item = await uploadResult.file.getItem();
      await item.update({
        Title: title,
        DocumentType: documentType,
        ProcessTypes: JSON.stringify(processTypes),
        Placeholders: JSON.stringify(placeholders),
        RequiresSignature: requiresSignature,
        IsActive: true
      });

      // Get the updated item with Id
      const updatedItem: any = await item.select('Id')();
      return await this.getTemplateById(updatedItem.Id);
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to create template:', error);
      throw error;
    }
  }

  /**
   * Update template
   */
  public async updateTemplate(
    templateId: number,
    updates: Partial<IJmlDocumentTemplate>
  ): Promise<void> {
    try {
      const metadata: any = {};

      if (updates.Title !== undefined) {
        metadata.Title = updates.Title;
      }
      if (updates.Description !== undefined) {
        metadata.Description = updates.Description;
      }
      if (updates.DocumentType !== undefined) {
        metadata.DocumentType = updates.DocumentType;
      }
      if (updates.ProcessTypes !== undefined) {
        metadata.ProcessTypes = JSON.stringify(updates.ProcessTypes);
      }
      if (updates.Placeholders !== undefined) {
        metadata.Placeholders = JSON.stringify(updates.Placeholders);
      }
      if (updates.RequiresSignature !== undefined) {
        metadata.RequiresSignature = updates.RequiresSignature;
      }
      if (updates.IsActive !== undefined) {
        metadata.IsActive = updates.IsActive;
      }

      await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .update(metadata);
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to update template:', error);
      throw error;
    }
  }

  /**
   * Delete template
   */
  public async deleteTemplate(templateId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .delete();
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to delete template:', error);
      throw error;
    }
  }

  /**
   * Get templates for process type
   */
  public async getTemplatesForProcess(processType: string, documentType?: DocumentType): Promise<IJmlDocumentTemplate[]> {
    try {
      const templates = await this.getTemplates(documentType);
      const matchingTemplates: IJmlDocumentTemplate[] = [];

      for (let i = 0; i < templates.length; i++) {
        const template = templates[i];
        if (template.ProcessTypes.indexOf(processType) !== -1) {
          matchingTemplates.push(template);
        }
      }

      return matchingTemplates;
    } catch (error) {
      logger.error('DocumentTemplateService', 'Failed to get templates for process:', error);
      return [];
    }
  }

  /**
   * Validate placeholder values
   */
  public validatePlaceholderValues(
    placeholders: ITemplatePlaceholder[],
    values: { [key: string]: string }
  ): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    for (let i = 0; i < placeholders.length; i++) {
      const placeholder = placeholders[i];

      if (placeholder.Required && !values[placeholder.Key]) {
        errors.push(`${placeholder.Label} is required`);
        continue;
      }

      const value = values[placeholder.Key];
      if (value && placeholder.ValidationPattern) {
        const regex = new RegExp(placeholder.ValidationPattern);
        if (!regex.test(value)) {
          errors.push(`${placeholder.Label} is invalid`);
        }
      }

      if (value && placeholder.DataType === 'date') {
        const date = new Date(value);
        if (isNaN(date.getTime())) {
          errors.push(`${placeholder.Label} must be a valid date`);
        }
      }

      if (value && placeholder.DataType === 'number') {
        if (isNaN(Number(value))) {
          errors.push(`${placeholder.Label} must be a number`);
        }
      }
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  /**
   * Map SharePoint item to template
   */
  private mapToTemplate(item: any): IJmlDocumentTemplate {
    return {
      Id: item.Id,
      Title: item.Title,
      Description: item.Description,
      DocumentType: item.DocumentType as DocumentType,
      TemplateUrl: item.File?.ServerRelativeUrl || '',
      ProcessTypes: item.ProcessTypes ? JSON.parse(item.ProcessTypes) : [],
      Placeholders: item.Placeholders ? JSON.parse(item.Placeholders) : [],
      RequiresSignature: item.RequiresSignature || false,
      SignatureProvider: item.SignatureProvider,
      IsActive: item.IsActive !== false,
      Created: new Date(item.Created),
      CreatedBy: {
        Id: item.AuthorId,
        Title: item.CreatedBy?.Title || item.Author?.Title || '',
        EMail: item.CreatedBy?.EMail || item.Author?.EMail || ''
      },
      Modified: new Date(item.Modified),
      ModifiedBy: {
        Id: item.EditorId,
        Title: item.ModifiedBy?.Title || item.Editor?.Title || '',
        EMail: item.ModifiedBy?.EMail || item.Editor?.EMail || ''
      }
    };
  }

  /**
   * Format date for template
   */
  private formatDate(date: Date): string {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const monthStr = month < 10 ? '0' + month : String(month);
    const dayStr = day < 10 ? '0' + day : String(day);
    return `${monthStr}/${dayStr}/${year}`;
  }
}
