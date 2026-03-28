// @ts-nocheck
import * as React from 'react';
import { IPolicyBulkUploadProps } from './IPolicyBulkUploadProps';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, IconButton, SearchBox, Dropdown, IDropdownOption,
  ProgressIndicator, Checkbox, TextField
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { StyledPanel } from '../../../components/StyledPanel';
import { PanelType } from '@fluentui/react';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { escapeHtml } from '../../../utils/sanitizeHtml';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';

// ============================================================================
// TYPES
// ============================================================================

type WizardStep = 1 | 2 | 3 | 4 | 5;
type ImportStatus = 'pending' | 'uploading' | 'uploaded' | 'classifying' | 'classified' | 'template-applied' | 'ready' | 'failed';
type MatchConfidence = 'Strong' | 'Likely' | 'Possible' | 'None';

interface IFileMetadata {
  title?: string;
  author?: string;
  subject?: string;
  category?: string;
  keywords?: string;
  company?: string;
  description?: string;
  created?: string;
  modified?: string;
}

interface IFastTrackTemplate {
  Id: number; Title: string; ProfileName: string; PolicyCategory: string;
  ComplianceRisk: string; ReadTimeframe: string; RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean; TargetDepartments: string;
}

interface IBulkImportItem {
  id: string;
  spId?: number;
  fileName: string;
  fileSize: number;
  fileType: string;
  file?: File;
  documentUrl?: string;
  status: ImportStatus;
  // Existing file metadata (extracted from document properties)
  existingMetadata?: IFileMetadata;
  hasExistingMetadata: boolean;
  useExistingMetadata: boolean; // user chose to skip AI classification
  // AI-extracted metadata
  extractedTitle?: string;
  suggestedCategory?: string;
  suggestedRisk?: string;
  suggestedDepartments?: string[];
  suggestedSummary?: string;
  suggestedKeyPoints?: string[];
  suggestedReadTimeframe?: string;
  // Template matching
  matchedTemplateId?: number;
  matchedTemplateName?: string;
  matchConfidence?: MatchConfidence;
  confirmedTemplateId?: number;
  templateApplied?: boolean;
  // User-edited fields
  confirmedTitle?: string;
  confirmedCategory?: string;
  confirmedRisk?: string;
  // Activity log
  error?: string;
}

interface IActivityLogEntry {
  time: Date;
  message: string;
  type: 'info' | 'success' | 'warning' | 'error';
}

interface IPolicyBulkUploadState {
  loading: boolean;
  detectedRole: PolicyManagerRole | null;
  wizardStep: WizardStep;
  completedSteps: Set<number>;
  imports: IBulkImportItem[];
  uploading: boolean;
  uploadProgress: number;
  classifying: boolean;
  classifyProgress: number;
  selectedIds: Set<string>;
  searchQuery: string;
  filterType: string;
  successMessage: string;
  errorMessage: string;
  dragOver: boolean;
  fastTrackTemplates: IFastTrackTemplate[];
  templatesLoaded: boolean;
  activityLog: IActivityLogEntry[];
  showBatchPanel: boolean;
  batchTemplateId: string;
  batchCategory: string;
  batchRisk: string;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const MAX_FILES = 50;
const MAX_FILE_SIZE = 25 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ['.docx', '.pdf', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.rtf', '.txt'];

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' },
  { key: 'IT Security', text: 'IT Security' }, { key: 'HR', text: 'Human Resources' },
  { key: 'Compliance', text: 'Compliance' }, { key: 'Data Protection', text: 'Data Protection' },
  { key: 'Health & Safety', text: 'Health & Safety' }, { key: 'Finance', text: 'Finance' },
  { key: 'Legal', text: 'Legal' }, { key: 'Operations', text: 'Operations' },
  { key: 'Governance', text: 'Governance' }, { key: 'Other', text: 'Other' }
];

const RISK_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' },
  { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' },
  { key: 'Informational', text: 'Informational' }
];

const WIZARD_STEPS = [
  { num: 1, title: 'Upload', desc: 'Add policy documents', icon: 'CloudUpload' },
  { num: 2, title: 'Review', desc: 'Check metadata & select', icon: 'ViewAll' },
  { num: 3, title: 'Classify', desc: 'AI-powered analysis', icon: 'Processing' },
  { num: 4, title: 'Templates', desc: 'Assign Fast Track', icon: 'Tag' },
  { num: 5, title: 'Finish', desc: 'Summary & next steps', icon: 'Accept' },
];

const F = "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif";

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyBulkUpload extends React.Component<IPolicyBulkUploadProps, IPolicyBulkUploadState> {
  private _isMounted = false;
  private _fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IPolicyBulkUploadProps) {
    super(props);
    this._fileInputRef = React.createRef();
    this.state = {
      loading: true, detectedRole: null,
      wizardStep: 1, completedSteps: new Set(),
      imports: [], uploading: false, uploadProgress: 0,
      classifying: false, classifyProgress: 0,
      selectedIds: new Set(), searchQuery: '', filterType: 'All',
      successMessage: '', errorMessage: '', dragOver: false,
      fastTrackTemplates: [], templatesLoaded: false,
      activityLog: [], showBatchPanel: false,
      batchTemplateId: '', batchCategory: '', batchRisk: ''
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this.detectRole();
  }
  public componentWillUnmount(): void { this._isMounted = false; }

  private async detectRole(): Promise<void> {
    try {
      const roleService = new RoleDetectionService(this.props.sp);
      const userRoles = await roleService.getCurrentUserRoles();
      const role = userRoles && userRoles.length > 0 ? getHighestPolicyRole(userRoles) : PolicyManagerRole.User;
      if (this._isMounted) this.setState({ detectedRole: role, loading: false });
    } catch {
      if (this._isMounted) this.setState({ detectedRole: PolicyManagerRole.Author, loading: false });
    }
  }

  private log(message: string, type: 'info' | 'success' | 'warning' | 'error' = 'info'): void {
    this.setState(prev => ({
      activityLog: [{ time: new Date(), message, type }, ...prev.activityLog]
    }));
  }

  // ============================================================================
  // FILE METADATA EXTRACTION
  // ============================================================================

  private async extractFileMetadata(file: File): Promise<IFileMetadata> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    const metadata: IFileMetadata = {};

    try {
      if (['docx', 'pptx', 'xlsx'].includes(ext)) {
        const arrayBuffer = await file.arrayBuffer();
        const bytes = new Uint8Array(arrayBuffer);
        const scanLimit = Math.min(bytes.length, 256000);
        let binaryStr = '';
        for (let i = 0; i < scanLimit; i++) binaryStr += String.fromCharCode(bytes[i]);

        // Extract from docProps/core.xml (Dublin Core metadata)
        const getXmlValue = (tag: string): string => {
          const patterns = [
            new RegExp(`<dc:${tag}[^>]*>([^<]+)</dc:${tag}>`, 'i'),
            new RegExp(`<cp:${tag}[^>]*>([^<]+)</cp:${tag}>`, 'i'),
            new RegExp(`<dcterms:${tag}[^>]*>([^<]+)</dcterms:${tag}>`, 'i'),
          ];
          for (const p of patterns) {
            const m = binaryStr.match(p);
            if (m) return m[1].trim();
          }
          return '';
        };

        metadata.title = getXmlValue('title');
        metadata.author = getXmlValue('creator') || getXmlValue('lastModifiedBy');
        metadata.subject = getXmlValue('subject');
        metadata.category = getXmlValue('category');
        metadata.keywords = getXmlValue('keywords');
        metadata.description = getXmlValue('description');
        metadata.created = getXmlValue('created');
        metadata.modified = getXmlValue('modified');

        // Extract Company from docProps/app.xml
        const companyMatch = binaryStr.match(/<Company>([^<]+)<\/Company>/i);
        if (companyMatch) metadata.company = companyMatch[1].trim();
      }

      if (ext === 'pdf') {
        const text = await file.slice(0, 4096).text();
        const getPdfInfo = (key: string): string => {
          const m = text.match(new RegExp(`/${key}\\s*\\(([^)]+)\\)`, 'i'));
          return m ? m[1].trim() : '';
        };
        metadata.title = getPdfInfo('Title');
        metadata.author = getPdfInfo('Author');
        metadata.subject = getPdfInfo('Subject');
        metadata.keywords = getPdfInfo('Keywords');
      }
    } catch (err) {
      console.warn(`[BulkUpload] Metadata extraction failed for ${file.name}:`, err);
    }

    return metadata;
  }

  // ============================================================================
  // DOCUMENT TEXT EXTRACTION (for AI classification)
  // ============================================================================

  private async extractTextFromFile(file: File): Promise<string> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    try {
      if (['txt', 'rtf', 'csv', 'md'].includes(ext)) {
        return (await file.text()).substring(0, 5000);
      }
      if (['docx', 'doc', 'pptx', 'xlsx'].includes(ext)) {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const limit = Math.min(bytes.length, 512000);
        let str = '';
        for (let i = 0; i < limit; i++) str += String.fromCharCode(bytes[i]);
        const parts: string[] = [];
        const regex = /<(?:w:|a:)?t[^>]*>([^<]+)<\/(?:w:|a:)?t>/g;
        let m: RegExpExecArray | null;
        while ((m = regex.exec(str)) !== null && parts.length < 800) {
          if (m[1].trim()) parts.push(m[1].trim());
        }
        if (parts.length > 5) return parts.join(' ').substring(0, 5000);
      }
      if (ext === 'pdf') {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const limit = Math.min(bytes.length, 512000);
        let text = '';
        for (let i = 0; i < limit; i++) {
          const c = bytes[i];
          text += (c >= 32 && c <= 126) || c === 10 || c === 13 ? String.fromCharCode(c) : ' ';
        }
        const sentences = text.replace(/\s{3,}/g, ' ').split(/\s{2,}/).filter(s => s.length > 20);
        if (sentences.length > 3) return sentences.join(' ').substring(0, 5000);
      }
    } catch { /* extraction failed */ }
    return `Document title: ${file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')}`;
  }

  // ============================================================================
  // STEP 1: FILE HANDLING
  // ============================================================================

  private handleFileDrop = (e: React.DragEvent): void => {
    e.preventDefault(); e.stopPropagation();
    this.setState({ dragOver: false });
    this.processFiles(Array.from(e.dataTransfer.files));
  };

  private handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.processFiles(Array.from(e.target.files || []));
    if (this._fileInputRef.current) this._fileInputRef.current.value = '';
  };

  private async processFiles(files: File[]): Promise<void> {
    const { imports } = this.state;
    const remaining = MAX_FILES - imports.length;
    if (remaining <= 0) { this.setState({ errorMessage: `Maximum ${MAX_FILES} files per batch.` }); return; }

    const newItems: IBulkImportItem[] = [];

    for (const file of files.slice(0, remaining)) {
      const ext = '.' + file.name.split('.').pop()?.toLowerCase();
      if (!ALLOWED_EXTENSIONS.includes(ext)) { this.log(`${file.name}: unsupported format`, 'warning'); continue; }
      if (file.size > MAX_FILE_SIZE) { this.log(`${file.name}: exceeds 25MB`, 'warning'); continue; }
      if (imports.some(i => i.fileName === file.name)) { this.log(`${file.name}: already added`, 'warning'); continue; }

      // Extract existing metadata from file properties
      const metadata = await this.extractFileMetadata(file);
      const hasMetadata = !!(metadata.title || metadata.category || metadata.keywords || metadata.subject);

      newItems.push({
        id: `imp_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        fileName: file.name,
        fileSize: file.size,
        fileType: ext,
        file,
        status: 'pending',
        existingMetadata: metadata,
        hasExistingMetadata: hasMetadata,
        useExistingMetadata: false,
        confirmedTitle: metadata.title || file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' '),
      });

      this.log(`Added: ${file.name}${hasMetadata ? ' (metadata found)' : ''}`, hasMetadata ? 'success' : 'info');
    }

    this.setState({ imports: [...imports, ...newItems] });
  }

  private removeImport = (id: string): void => {
    const item = this.state.imports.find(i => i.id === id);
    if (item) this.log(`Removed: ${item.fileName}`, 'info');
    this.setState(prev => ({
      imports: prev.imports.filter(i => i.id !== id),
      selectedIds: (() => { const s = new Set(prev.selectedIds); s.delete(id); return s; })()
    }));
  };

  // ============================================================================
  // STEP 1: UPLOAD TO SHAREPOINT
  // ============================================================================

  private async uploadToSharePoint(): Promise<void> {
    const { imports } = this.state;
    const toUpload = imports.filter(i => i.status === 'pending' && i.file);
    if (toUpload.length === 0) return;

    this.setState({ uploading: true, uploadProgress: 0 });
    this.log(`Uploading ${toUpload.length} file${toUpload.length !== 1 ? 's' : ''}...`, 'info');

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
    const siteRelUrl = this.props.context?.pageContext?.web?.serverRelativeUrl || '/sites/PolicyManager';

    // Get digest
    let digest = '';
    try {
      const { SPHttpClient } = await import('@microsoft/sp-http');
      const resp = await this.props.context.spHttpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {});
      const d = await resp.json();
      digest = d?.FormDigestValue || d?.d?.GetContextWebInformation?.FormDigestValue || '';
    } catch { digest = (document.getElementById('__REQUESTDIGEST') as HTMLInputElement)?.value || ''; }

    // Ensure folder
    const folderUrl = `${siteRelUrl}/${PM_LISTS.POLICY_SOURCE_DOCUMENTS}/BulkImports`;
    try {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', `${siteUrl}/_api/web/folders`, false);
      xhr.setRequestHeader('Accept', 'application/json; odata=verbose');
      xhr.setRequestHeader('Content-Type', 'application/json; odata=verbose');
      xhr.setRequestHeader('X-RequestDigest', digest);
      xhr.send(JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': folderUrl }));
    } catch { /* exists */ }

    let processed = 0;
    let succeeded = 0;

    for (const item of toUpload) {
      try {
        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, status: 'uploading' as ImportStatus } : i) });

        const buf = await item.file.arrayBuffer();
        const safeName = item.fileName.replace(/[#%&*:<>?\/\\{|}~]/g, '_');
        const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folderUrl}')/Files/AddUsingPath(decodedurl='${encodeURIComponent(safeName)}',overwrite=true)`;

        const docUrl: string = await new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.open('POST', endpoint, true);
          xhr.setRequestHeader('Accept', 'application/json; odata=verbose');
          xhr.setRequestHeader('Content-Type', 'application/octet-stream');
          xhr.setRequestHeader('X-RequestDigest', digest);
          xhr.responseType = 'json';
          xhr.onload = () => xhr.status >= 200 && xhr.status < 300 ? resolve(xhr.response?.d?.ServerRelativeUrl || '') : reject(new Error(`${xhr.status}`));
          xhr.onerror = () => reject(new Error('Network error'));
          xhr.send(new Uint8Array(buf));
        });

        // Create policy stub
        const title = item.confirmedTitle || item.fileName.replace(/\.[^.]+$/, '');
        const result = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.add({ Title: title, PolicyStatus: 'Draft' });
        const spId = result?.data?.Id || result?.data?.id;
        if (spId && docUrl) {
          try { await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(spId).update({ PolicyName: title, DocumentURL: docUrl }); } catch { /* */ }
        }

        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, spId, documentUrl: docUrl, status: 'uploaded' as ImportStatus } : i) });
        this.log(`Uploaded: ${item.fileName}`, 'success');
        succeeded++;
      } catch (err) {
        console.error(`[BulkUpload] Upload failed: ${item.fileName}`, err);
        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, status: 'failed' as ImportStatus, error: String(err) } : i) });
        this.log(`Failed: ${item.fileName} — ${(err as Error).message}`, 'error');
      }
      processed++;
      if (this._isMounted) this.setState({ uploadProgress: Math.round((processed / toUpload.length) * 100) });
    }

    const completed = new Set(this.state.completedSteps);
    completed.add(1);
    this.setState({ uploading: false, completedSteps: completed, wizardStep: 2 });
    this.log(`Upload complete: ${succeeded}/${toUpload.length} succeeded`, succeeded === toUpload.length ? 'success' : 'warning');
  }

  // ============================================================================
  // STEP 3: AI CLASSIFICATION
  // ============================================================================

  private async classifySelected(): Promise<void> {
    const { imports, selectedIds } = this.state;
    const toClassify = imports.filter(i => selectedIds.has(i.id) && i.status === 'uploaded' && !i.useExistingMetadata);
    if (toClassify.length === 0) { this.setState({ errorMessage: 'No items selected for classification.' }); return; }

    this.setState({ classifying: true, classifyProgress: 0 });
    this.log(`Classifying ${toClassify.length} file${toClassify.length !== 1 ? 's' : ''} with AI...`, 'info');
    await this.loadFastTrackTemplates();

    let functionUrl = '';
    try {
      const c = await this.props.sp.web.lists.getByTitle('PM_Configuration').items.filter("ConfigKey eq 'Integration.AI.Chat.FunctionUrl'").select('ConfigValue').top(1)();
      functionUrl = c[0]?.ConfigValue || '';
    } catch { /* */ }
    if (!functionUrl) functionUrl = localStorage.getItem('PM_AI_ChatFunctionUrl') || '';

    let processed = 0;
    for (const item of toClassify) {
      try {
        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, status: 'classifying' as ImportStatus } : i) });

        let docContent = '';
        if (item.file) { try { docContent = await this.extractTextFromFile(item.file); } catch { /* */ } }

        let suggestions: any = {};
        if (functionUrl) {
          const contentCtx = docContent.length > 100 ? `\n\nDocument content:\n"""${docContent.substring(0, 3000)}"""` : '';
          try {
            const ctrl = new AbortController();
            const t = setTimeout(() => ctrl.abort(), 45000);
            const resp = await fetch(functionUrl, {
              method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ mode: 'author-assist', message: `Classify this policy document. Filename: "${item.fileName}"${contentCtx}\n\nExtract:\n1. Title (clean professional title)\n2. PolicyCategory (IT Security, HR, Compliance, Data Protection, Health & Safety, Finance, Legal, Operations, Governance, Other)\n3. ComplianceRisk (Critical, High, Medium, Low, Informational)\n4. Departments (comma-separated)\n5. Summary (1-2 sentences)\n6. KeyPoints (3-5 items)\n7. ReadTimeframe (Immediate, Day 1, Day 3, Week 1, Week 2, Month 1)\n8. RequiresAcknowledgement (true/false)\n\nRespond ONLY with JSON: {title, category, risk, departments, summary, keyPoints, readTimeframe, requiresAck}`, history: [], context: [] }),
              signal: ctrl.signal
            });
            clearTimeout(t);
            if (resp.ok) {
              const data = await resp.json();
              const content = data?.response || data?.content || '';
              try { const m = content.match(/\{[\s\S]*\}/); if (m) suggestions = JSON.parse(m[0]); } catch { /* */ }
            }
          } catch { /* AI failed */ }
        }

        if (!suggestions.category) suggestions = this.heuristicClassify(item.fileName, item.confirmedTitle || '');

        // Build classified item
        const classified: Partial<IBulkImportItem> = {
          status: 'classified' as ImportStatus,
          extractedTitle: suggestions.title || item.confirmedTitle,
          suggestedCategory: suggestions.category || 'Other',
          suggestedRisk: suggestions.risk || 'Medium',
          suggestedDepartments: Array.isArray(suggestions.departments) ? suggestions.departments : (suggestions.departments || '').split(',').map((d: string) => d.trim()).filter(Boolean),
          suggestedSummary: suggestions.summary || '',
          suggestedKeyPoints: Array.isArray(suggestions.keyPoints) ? suggestions.keyPoints : [],
          suggestedReadTimeframe: suggestions.readTimeframe || 'Week 1',
          confirmedTitle: suggestions.title || item.confirmedTitle,
          confirmedCategory: suggestions.category || 'Other',
          confirmedRisk: suggestions.risk || 'Medium',
        };

        // Match template
        const tempItem = { ...item, ...classified } as IBulkImportItem;
        const match = this.matchTemplate(tempItem);
        if (match) { classified.matchedTemplateId = match.templateId; classified.matchedTemplateName = match.templateName; classified.matchConfidence = match.confidence; classified.confirmedTemplateId = match.templateId; }

        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, ...classified, file: undefined } : i) });
        this.log(`Classified: ${item.fileName} → ${classified.suggestedCategory} (${classified.suggestedRisk})`, 'success');
      } catch (err) {
        this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, status: 'uploaded' as ImportStatus, error: String(err) } : i) });
        this.log(`Classification failed: ${item.fileName}`, 'error');
      }
      processed++;
      if (this._isMounted) this.setState({ classifyProgress: Math.round((processed / toClassify.length) * 100) });
    }

    const completed = new Set(this.state.completedSteps);
    completed.add(3);
    this.setState({ classifying: false, completedSteps: completed });
    this.log(`Classification complete: ${processed} processed`, 'success');
  }

  private heuristicClassify(fileName: string, title: string): any {
    const text = (fileName + ' ' + title).toLowerCase();
    let category = 'Other', risk = 'Medium';
    if (/security|cyber|access|password|firewall/i.test(text)) { category = 'IT Security'; risk = 'High'; }
    else if (/hr|human|employee|leave|conduct|harassment/i.test(text)) { category = 'HR'; risk = 'Medium'; }
    else if (/compliance|regulat|audit|sox|iso/i.test(text)) { category = 'Compliance'; risk = 'High'; }
    else if (/data|privacy|gdpr|pii|personal/i.test(text)) { category = 'Data Protection'; risk = 'Critical'; }
    else if (/health|safety|incident|hazard/i.test(text)) { category = 'Health & Safety'; risk = 'High'; }
    else if (/financ|expense|procurement|budget|travel/i.test(text)) { category = 'Finance'; risk = 'Medium'; }
    else if (/legal|contract|nda|confidential/i.test(text)) { category = 'Legal'; risk = 'High'; }
    else if (/operat|process|procedure|standard/i.test(text)) { category = 'Operations'; risk = 'Low'; }
    else if (/govern|board|ethic|whistleblow/i.test(text)) { category = 'Governance'; risk = 'High'; }
    const cleanTitle = (fileName).replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    return { title: cleanTitle, category, risk, departments: ['All Employees'], summary: '', keyPoints: [], readTimeframe: 'Week 1', requiresAck: ['Critical', 'High'].includes(risk) };
  }

  // ============================================================================
  // TEMPLATE MATCHING
  // ============================================================================

  private async loadFastTrackTemplates(): Promise<void> {
    if (this.state.templatesLoaded) return;
    try {
      let items: any[] = [];
      try {
        items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
          .items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk', 'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz', 'TargetDepartments', 'IsActive')
          .orderBy('Title').top(100)();
      } catch {
        try { items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES).items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk').orderBy('Title').top(100)(); } catch { /* */ }
      }
      const templates = items.filter((t: any) => t.IsActive !== false).map((t: any) => ({
        Id: t.Id, Title: t.Title || t.ProfileName || `Template ${t.Id}`, ProfileName: t.ProfileName || t.Title || '',
        PolicyCategory: t.PolicyCategory || '', ComplianceRisk: t.ComplianceRisk || 'Medium',
        ReadTimeframe: t.ReadTimeframe || 'Week 1', RequiresAcknowledgement: t.RequiresAcknowledgement !== false,
        RequiresQuiz: t.RequiresQuiz || false, TargetDepartments: t.TargetDepartments || ''
      }));
      if (this._isMounted) this.setState({ fastTrackTemplates: templates, templatesLoaded: true });
    } catch { if (this._isMounted) this.setState({ templatesLoaded: true }); }
  }

  private matchTemplate(item: IBulkImportItem): { templateId: number; templateName: string; confidence: MatchConfidence } | null {
    const { fastTrackTemplates } = this.state;
    if (fastTrackTemplates.length === 0) return null;
    const cat = (item.suggestedCategory || '').toLowerCase();
    const risk = (item.suggestedRisk || '').toLowerCase();
    let best: IFastTrackTemplate | null = null, bestScore = 0;
    for (const t of fastTrackTemplates) {
      let score = 0;
      if (t.PolicyCategory.toLowerCase() === cat) score += 3;
      if (t.ComplianceRisk.toLowerCase() === risk) score += 2;
      if (score > bestScore) { bestScore = score; best = t; }
    }
    if (!best || bestScore === 0) return null;
    return { templateId: best.Id, templateName: best.Title, confidence: bestScore >= 5 ? 'Strong' : bestScore >= 3 ? 'Likely' : 'Possible' };
  }

  private async applyTemplateToItem(itemId: string, templateId: number): Promise<void> {
    const template = this.state.fastTrackTemplates.find(t => t.Id === templateId);
    if (!template) return;
    const item = this.state.imports.find(i => i.id === itemId);
    if (!item?.spId) return;
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({
        PolicyCategory: template.PolicyCategory, ComplianceRisk: template.ComplianceRisk,
        ReadTimeframe: template.ReadTimeframe, RequiresAcknowledgement: template.RequiresAcknowledgement,
        RequiresQuiz: template.RequiresQuiz, Departments: template.TargetDepartments
      });
    } catch { /* */ }
    this.setState({ imports: this.state.imports.map(i => i.id === itemId ? { ...i, confirmedCategory: template.PolicyCategory, confirmedRisk: template.ComplianceRisk, confirmedTemplateId: templateId, templateApplied: true, status: 'template-applied' as ImportStatus } : i) });
    this.log(`Template applied: ${template.Title} → ${item.fileName}`, 'success');
  }

  // ============================================================================
  // RENDER: MAIN
  // ============================================================================

  public render(): React.ReactElement {
    const { detectedRole, loading } = this.state;
    if (loading) return (<ErrorBoundary fallbackMessage="An error occurred in Bulk Upload."><JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}><div style={{ padding: 60, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading..." /></div></JmlAppLayout></ErrorBoundary>);
    if (detectedRole !== null && !hasMinimumRole(detectedRole, PolicyManagerRole.Author)) return (<ErrorBoundary fallbackMessage="An error occurred."><JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}><section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}><Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} /><Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>Access Denied</Text></section></JmlAppLayout></ErrorBoundary>);

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Bulk Upload.">
        <JmlAppLayout title={this.props.title || 'Policy Bulk Upload'} context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}>
          {this.renderWizard()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ============================================================================
  // RENDER: WIZARD CHROME
  // ============================================================================

  private renderWizard(): React.ReactElement {
    const { wizardStep, completedSteps, imports, successMessage, errorMessage } = this.state;
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    return (
      <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr', minHeight: 'calc(100vh - 180px)', background: '#fff', borderRadius: 10, overflow: 'hidden', border: '1px solid #e2e8f0', margin: '0 auto', maxWidth: 1400 }}>
        {/* Sidebar */}
        <aside style={{ background: '#fff', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column' }}>
          <div style={{ padding: '24px 20px 16px', borderBottom: '1px solid #e2e8f0' }}>
            <Text style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', display: 'block' }}>Bulk Upload</Text>
            <Text style={{ fontSize: 11, color: '#94a3b8', marginTop: 2, display: 'block' }}>5 steps to import policies</Text>
          </div>
          <div style={{ flex: 1, padding: '8px 0' }}>
            {WIZARD_STEPS.map(step => {
              const done = completedSteps.has(step.num);
              const active = step.num === wizardStep;
              const clickable = done || step.num <= wizardStep;
              return (
                <div key={step.num}
                  onClick={() => clickable && this.setState({ wizardStep: step.num as WizardStep })}
                  style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '12px 20px', cursor: clickable ? 'pointer' : 'default', borderLeft: active ? '3px solid #0d9488' : '3px solid transparent', background: active ? '#f0fdfa' : 'transparent', transition: 'all 0.15s' }}>
                  <div style={{ width: 28, height: 28, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0, background: done ? '#0d9488' : active ? '#f0fdfa' : '#fff', color: done ? '#fff' : active ? '#0d9488' : '#94a3b8', border: `2px solid ${done ? '#0d9488' : active ? '#0d9488' : '#e2e8f0'}` }}>
                    {done ? <Icon iconName="CheckMark" style={{ fontSize: 11 }} /> : step.num}
                  </div>
                  <div>
                    <div style={{ fontWeight: active ? 600 : 500, color: active ? '#0d9488' : done ? '#0f172a' : '#475569', fontSize: 13 }}>{step.title}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8' }}>{step.desc}</div>
                  </div>
                </div>
              );
            })}
          </div>
          {/* KPI summary at bottom of sidebar */}
          <div style={{ padding: '16px 20px', borderTop: '1px solid #e2e8f0' }}>
            {[
              { label: 'Total', val: imports.length, color: '#475569' },
              { label: 'Uploaded', val: imports.filter(i => !['pending', 'failed'].includes(i.status)).length, color: '#2563eb' },
              { label: 'Classified', val: imports.filter(i => ['classified', 'template-applied', 'ready'].includes(i.status)).length, color: '#7c3aed' },
              { label: 'Templates', val: imports.filter(i => i.templateApplied).length, color: '#059669' },
            ].map(k => (
              <div key={k.label} style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 0', fontSize: 12 }}>
                <span style={{ color: '#64748b' }}>{k.label}</span>
                <span style={{ fontWeight: 700, color: k.color }}>{k.val}</span>
              </div>
            ))}
          </div>
        </aside>

        {/* Main content */}
        <div style={{ padding: '28px 36px', overflowY: 'auto', background: '#f8fafc' }}>
          {successMessage && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ successMessage: '' })} style={{ marginBottom: 12 }}>{successMessage}</MessageBar>}
          {errorMessage && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ errorMessage: '' })} style={{ marginBottom: 12 }}>{errorMessage}</MessageBar>}

          {wizardStep === 1 && this.renderStep1_Upload()}
          {wizardStep === 2 && this.renderStep2_Review()}
          {wizardStep === 3 && this.renderStep3_Classify()}
          {wizardStep === 4 && this.renderStep4_Templates(siteUrl)}
          {wizardStep === 5 && this.renderStep5_Finish(siteUrl)}

          {/* Navigation footer */}
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 24, paddingTop: 16, borderTop: '1px solid #e2e8f0' }}>
            <DefaultButton text="Previous" iconProps={{ iconName: 'ChevronLeft' }} disabled={wizardStep === 1}
              onClick={() => this.setState({ wizardStep: Math.max(1, wizardStep - 1) as WizardStep })}
              styles={{ root: { borderRadius: 4, visibility: wizardStep === 1 ? 'hidden' : 'visible' } }} />
            {wizardStep < 5 ? (
              <PrimaryButton text="Next" onClick={() => this.handleNext()}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}>
                Next <Icon iconName="ChevronRight" style={{ marginLeft: 6 }} />
              </PrimaryButton>
            ) : null}
          </div>
        </div>
      </div>
    );
  }

  private handleNext(): void {
    const { wizardStep, imports } = this.state;
    const completed = new Set(this.state.completedSteps);
    completed.add(wizardStep);
    const nextStep = Math.min(5, wizardStep + 1) as WizardStep;

    // Auto-select all uploaded items when entering Step 3
    if (nextStep === 3) {
      const uploadedIds = new Set(imports.filter(i => i.status === 'uploaded' && !i.useExistingMetadata).map(i => i.id));
      this.setState({ wizardStep: nextStep, completedSteps: completed, selectedIds: uploadedIds });
    } else if (nextStep === 4) {
      this.loadFastTrackTemplates();
      this.setState({ wizardStep: nextStep, completedSteps: completed });
    } else {
      this.setState({ wizardStep: nextStep, completedSteps: completed });
    }
  }

  // ============================================================================
  // STEP 1: UPLOAD
  // ============================================================================

  private renderStep1_Upload(): React.ReactElement {
    const { imports, uploading, uploadProgress, dragOver } = this.state;
    const pendingFiles = imports.filter(i => i.status === 'pending');

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Upload Policy Documents</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 20px' }}>Drag and drop or browse for files. Supported: DOCX, PDF, XLSX, PPTX (up to 25MB each, max {MAX_FILES} files)</p>

        {/* Drop zone */}
        <div onDragOver={(e) => { e.preventDefault(); this.setState({ dragOver: true }); }} onDragLeave={() => this.setState({ dragOver: false })} onDrop={this.handleFileDrop}
          onClick={() => this._fileInputRef.current?.click()}
          style={{ border: `2px dashed ${dragOver ? '#0d9488' : '#cbd5e1'}`, borderRadius: 10, padding: '40px 32px', textAlign: 'center', cursor: 'pointer', background: dragOver ? '#f0fdfa' : '#fff', transition: 'all 0.2s', marginBottom: 20 }}>
          <input ref={this._fileInputRef} type="file" multiple accept={ALLOWED_EXTENSIONS.join(',')} onChange={this.handleFileSelect} style={{ display: 'none' }} />
          <svg viewBox="0 0 24 24" fill="none" width="36" height="36" style={{ color: dragOver ? '#0d9488' : '#94a3b8', marginBottom: 10 }}><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /><path d="M17 8l-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /><path d="M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /></svg>
          <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', marginBottom: 4 }}>{dragOver ? 'Drop files here' : 'Drag & drop policy documents here'}</div>
          <div style={{ fontSize: 12, color: '#94a3b8' }}>or click to browse</div>
        </div>

        {/* File list */}
        {imports.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 80px 80px 90px 32px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
              <div>File</div><div>Size</div><div>Type</div><div>Metadata</div><div></div>
            </div>
            {imports.map(item => {
              const sizeStr = item.fileSize < 1024 * 1024 ? `${Math.round(item.fileSize / 1024)} KB` : `${(item.fileSize / (1024 * 1024)).toFixed(1)} MB`;
              return (
                <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '1fr 80px 80px 90px 32px', padding: '10px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: 13 }}>
                  <div>
                    <div style={{ fontWeight: 600, color: '#0f172a' }}>{item.confirmedTitle || item.fileName}</div>
                    <div style={{ fontSize: 11, color: '#94a3b8' }}>{item.fileName}</div>
                  </div>
                  <div style={{ color: '#64748b', fontSize: 12 }}>{sizeStr}</div>
                  <div><span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0f9ff', color: '#0369a1' }}>{item.fileType.replace('.', '').toUpperCase()}</span></div>
                  <div>
                    {item.hasExistingMetadata ? (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Found</span>
                    ) : (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#fef3c7', color: '#d97706' }}>None</span>
                    )}
                  </div>
                  <div><IconButton iconProps={{ iconName: 'Cancel' }} title="Remove" onClick={() => this.removeImport(item.id)} styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 11, color: '#dc2626' } }} /></div>
                </div>
              );
            })}
          </div>
        )}

        {uploading && <ProgressIndicator label={`Uploading... ${uploadProgress}%`} percentComplete={uploadProgress / 100} style={{ marginBottom: 16 }} />}

        {pendingFiles.length > 0 && !uploading && (
          <PrimaryButton text={`Upload ${pendingFiles.length} File${pendingFiles.length !== 1 ? 's' : ''}`} iconProps={{ iconName: 'CloudUpload' }}
            onClick={() => this.uploadToSharePoint()}
            styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
        )}
      </>
    );
  }

  // ============================================================================
  // STEP 2: REVIEW
  // ============================================================================

  private renderStep2_Review(): React.ReactElement {
    const { imports, searchQuery, filterType, selectedIds } = this.state;
    const uploaded = imports.filter(i => !['pending', 'failed'].includes(i.status));

    let filtered = uploaded;
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); filtered = filtered.filter(i => (i.confirmedTitle || i.fileName).toLowerCase().includes(q)); }
    if (filterType !== 'All') filtered = filtered.filter(i => i.fileType === `.${filterType.toLowerCase()}`);

    const allSelected = filtered.length > 0 && filtered.every(i => selectedIds.has(i.id));
    const withMetadata = filtered.filter(i => i.hasExistingMetadata);

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Review Uploaded Documents</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 16px' }}>Review existing metadata, edit titles, and select which files to classify with AI. Files with good metadata can skip classification.</p>

        {withMetadata.length > 0 && (
          <MessageBar messageBarType={MessageBarType.success} style={{ marginBottom: 12 }}>
            {withMetadata.length} file{withMetadata.length !== 1 ? 's have' : ' has'} existing metadata. You can accept this metadata and skip AI classification for these files.
          </MessageBar>
        )}

        {/* Toolbar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12, flexWrap: 'wrap' }}>
          <SearchBox placeholder="Search..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 220 } }} />
          <Dropdown selectedKey={filterType} options={[{ key: 'All', text: 'All Types' }, { key: 'DOCX', text: 'DOCX' }, { key: 'PDF', text: 'PDF' }, { key: 'XLSX', text: 'XLSX' }, { key: 'PPTX', text: 'PPTX' }]}
            onChange={(_, opt) => this.setState({ filterType: String(opt?.key || 'All') })}
            styles={{ root: { width: 120 }, title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: '#64748b' }}>{selectedIds.size} selected</span>
        </div>

        {/* Table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '36px 1fr 120px 100px 100px 100px 80px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" checked={allSelected} onChange={() => {
              if (allSelected) this.setState({ selectedIds: new Set() }); else this.setState({ selectedIds: new Set(filtered.map(i => i.id)) });
            }} /></div>
            <div>Title / File</div><div>Existing Meta</div><div>Author</div><div>Category</div><div>Keywords</div><div>Skip AI</div>
          </div>
          {filtered.map(item => {
            const isSelected = selectedIds.has(item.id);
            const meta = item.existingMetadata || {};
            return (
              <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '36px 1fr 120px 100px 100px 100px 80px', padding: '10px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: isSelected ? '#f0fdfa' : item.useExistingMetadata ? '#f0fdf4' : '#fff' }}>
                <div><input type="checkbox" checked={isSelected} onChange={() => {
                  const next = new Set(selectedIds); if (next.has(item.id)) next.delete(item.id); else next.add(item.id); this.setState({ selectedIds: next });
                }} /></div>
                <div>
                  <input type="text" value={item.confirmedTitle || ''} onChange={(e) => {
                    const v = (e.target as HTMLInputElement).value;
                    this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, confirmedTitle: v } : i) });
                  }} style={{ width: '100%', border: 'none', background: 'transparent', fontSize: 13, fontWeight: 600, color: '#0f172a', outline: 'none', padding: '2px 0' }} />
                  <div style={{ fontSize: 10, color: '#94a3b8' }}>{item.fileName}</div>
                </div>
                <div>{item.hasExistingMetadata ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Found</span> : <span style={{ fontSize: 10, color: '#cbd5e1' }}>—</span>}</div>
                <div style={{ fontSize: 11, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{meta.author || '—'}</div>
                <div style={{ fontSize: 11, color: '#475569' }}>{meta.category || '—'}</div>
                <div style={{ fontSize: 11, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{meta.keywords || '—'}</div>
                <div>
                  {item.hasExistingMetadata && (
                    <Checkbox checked={item.useExistingMetadata} onChange={(_, checked) => {
                      this.setState({ imports: this.state.imports.map(i => i.id === item.id ? { ...i, useExistingMetadata: !!checked } : i) });
                    }} styles={{ root: { marginBottom: 0 } }} />
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // ============================================================================
  // STEP 3: CLASSIFY
  // ============================================================================

  private renderStep3_Classify(): React.ReactElement {
    const { imports, classifying, classifyProgress, selectedIds } = this.state;
    const needsClassification = imports.filter(i => i.status === 'uploaded' && !i.useExistingMetadata);
    const classified = imports.filter(i => ['classified', 'template-applied'].includes(i.status));
    const skipped = imports.filter(i => i.useExistingMetadata);

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>AI Classification</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 16px' }}>
          {needsClassification.length} file{needsClassification.length !== 1 ? 's' : ''} ready for AI classification.
          {skipped.length > 0 && ` ${skipped.length} skipped (using existing metadata).`}
        </p>

        {classifying && <ProgressIndicator label={`Classifying... ${classifyProgress}%`} percentComplete={classifyProgress / 100} style={{ marginBottom: 16 }} />}

        {!classifying && needsClassification.length > 0 && classified.length === 0 && (
          <div style={{ display: 'flex', gap: 8, marginBottom: 20 }}>
            <PrimaryButton text={`Classify ${selectedIds.size > 0 ? selectedIds.size : needsClassification.length} File${(selectedIds.size > 0 ? selectedIds.size : needsClassification.length) !== 1 ? 's' : ''}`}
              iconProps={{ iconName: 'Processing' }} onClick={() => {
                if (selectedIds.size === 0) this.setState({ selectedIds: new Set(needsClassification.map(i => i.id)) }, () => this.classifySelected());
                else this.classifySelected();
              }}
              styles={{ root: { background: '#7c3aed', borderColor: '#7c3aed', borderRadius: 4 }, rootHovered: { background: '#6d28d9', borderColor: '#6d28d9' } }} />
          </div>
        )}

        {/* Classification results */}
        {classified.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: 12 }}>
            {classified.map(item => (
              <div key={item.id} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16, borderLeft: `4px solid ${item.suggestedRisk === 'Critical' ? '#dc2626' : item.suggestedRisk === 'High' ? '#d97706' : '#0d9488'}` }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', marginBottom: 4 }}>{item.extractedTitle || item.confirmedTitle}</div>
                <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 10 }}>{item.fileName}</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 10 }}>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: '#f5f3ff', color: '#7c3aed' }}>{item.suggestedCategory}</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: item.suggestedRisk === 'Critical' ? '#fef2f2' : item.suggestedRisk === 'High' ? '#fff7ed' : '#f0fdf4', color: item.suggestedRisk === 'Critical' ? '#dc2626' : item.suggestedRisk === 'High' ? '#d97706' : '#059669' }}>{item.suggestedRisk}</span>
                  {item.matchConfidence && <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: '#eff6ff', color: '#2563eb' }}>{item.matchConfidence} match</span>}
                </div>
                {item.suggestedSummary && <div style={{ fontSize: 12, color: '#64748b', lineHeight: 1.5, marginBottom: 8 }}>{item.suggestedSummary.substring(0, 120)}{item.suggestedSummary.length > 120 ? '...' : ''}</div>}
                <div style={{ fontSize: 11, color: '#94a3b8' }}>Departments: {(item.suggestedDepartments || []).join(', ') || 'All'}</div>
              </div>
            ))}
          </div>
        )}

        {needsClassification.length === 0 && classified.length === 0 && (
          <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8' }}>All files are either classified or using existing metadata. Proceed to Templates.</div>
        )}
      </>
    );
  }

  // ============================================================================
  // STEP 4: TEMPLATES
  // ============================================================================

  private renderStep4_Templates(siteUrl: string): React.ReactElement {
    const { imports, fastTrackTemplates, showBatchPanel, selectedIds } = this.state;
    const classifiedItems = imports.filter(i => ['classified', 'template-applied'].includes(i.status));
    const templateOptions: IDropdownOption[] = [{ key: '', text: '— No template —' }, ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))];

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Assign Fast Track Templates</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 16px' }}>Match classified policies to Fast Track Templates for automatic metadata pre-fill. AI has suggested matches where possible.</p>

        <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
          <DefaultButton text="Auto-Match All" iconProps={{ iconName: 'Processing' }}
            onClick={async () => { for (const item of classifiedItems) { if (item.matchedTemplateId && !item.templateApplied) await this.applyTemplateToItem(item.id, item.matchedTemplateId); } }}
            styles={{ root: { borderRadius: 4, color: '#059669', borderColor: '#bbf7d0' }, rootHovered: { background: '#f0fdf4' } }} />
          {selectedIds.size > 0 && <DefaultButton text={`Batch Assign (${selectedIds.size})`} iconProps={{ iconName: 'Tag' }} onClick={() => this.setState({ showBatchPanel: true })} styles={{ root: { borderRadius: 4 } }} />}
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '36px 1fr 120px 90px 200px 80px 60px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" onChange={(e) => { if ((e.target as HTMLInputElement).checked) this.setState({ selectedIds: new Set(classifiedItems.map(i => i.id)) }); else this.setState({ selectedIds: new Set() }); }} /></div>
            <div>Policy</div><div>Category</div><div>Risk</div><div>Fast Track Template</div><div>Match</div><div>Actions</div>
          </div>
          {classifiedItems.map(item => (
            <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '36px 1fr 120px 90px 200px 80px 60px', padding: '10px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: item.templateApplied ? '#f0fdf4' : '#fff' }}>
              <div><input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => { const next = new Set(selectedIds); if (next.has(item.id)) next.delete(item.id); else next.add(item.id); this.setState({ selectedIds: next }); }} /></div>
              <div><div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{item.confirmedTitle}</div><div style={{ fontSize: 10, color: '#94a3b8' }}>{item.fileName}</div></div>
              <div style={{ fontSize: 12, color: '#475569' }}>{item.confirmedCategory || item.suggestedCategory}</div>
              <div style={{ fontSize: 12, color: '#475569' }}>{item.confirmedRisk || item.suggestedRisk}</div>
              <div>
                <Dropdown selectedKey={item.confirmedTemplateId ? String(item.confirmedTemplateId) : (item.matchedTemplateId ? String(item.matchedTemplateId) : '')} options={templateOptions}
                  onChange={(_, opt) => { if (opt?.key) this.applyTemplateToItem(item.id, parseInt(String(opt.key))); }}
                  styles={{ root: { minWidth: 0 }, title: { fontSize: 12, height: 28, lineHeight: '26px', borderRadius: 4, borderColor: item.templateApplied ? '#bbf7d0' : '#e2e8f0' }, caretDownWrapper: { height: 28, lineHeight: '28px' } }} />
              </div>
              <div>{item.matchConfidence ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: item.matchConfidence === 'Strong' ? '#f0fdf4' : '#eff6ff', color: item.matchConfidence === 'Strong' ? '#059669' : '#2563eb' }}>{item.matchConfidence}</span> : item.templateApplied ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Applied</span> : <span style={{ color: '#cbd5e1', fontSize: 10 }}>—</span>}</div>
              <div>{item.spId && <IconButton iconProps={{ iconName: 'Edit' }} title="Open in Policy Builder" href={`${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${item.spId}`} styles={{ root: { width: 26, height: 26 }, icon: { fontSize: 12, color: '#0d9488' } }} />}</div>
            </div>
          ))}
        </div>

        {this.renderBatchPanel()}
      </>
    );
  }

  // ============================================================================
  // STEP 5: FINISH
  // ============================================================================

  private renderStep5_Finish(siteUrl: string): React.ReactElement {
    const { imports, activityLog } = this.state;
    const uploaded = imports.filter(i => !['pending', 'failed'].includes(i.status)).length;
    const classified = imports.filter(i => ['classified', 'template-applied', 'ready'].includes(i.status)).length;
    const templated = imports.filter(i => i.templateApplied).length;
    const failed = imports.filter(i => i.status === 'failed').length;

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Import Summary</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 20px' }}>Your policies have been imported as drafts. Open them in the Policy Builder to add content and submit for review.</p>

        {/* KPI cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
          {[
            { label: 'Uploaded', value: uploaded, color: '#2563eb' },
            { label: 'Classified', value: classified, color: '#7c3aed' },
            { label: 'Templates Applied', value: templated, color: '#059669' },
            { label: 'Failed', value: failed, color: failed > 0 ? '#dc2626' : '#94a3b8' },
          ].map(k => (
            <div key={k.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.color}`, padding: '16px 18px' }}>
              <div style={{ fontSize: 26, fontWeight: 700, color: k.color }}>{k.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.label}</div>
            </div>
          ))}
        </div>

        <div style={{ display: 'flex', gap: 8, marginBottom: 24 }}>
          <PrimaryButton text="Open Drafts & Pipeline" iconProps={{ iconName: 'ViewAll' }} href={`${siteUrl}/SitePages/PolicyAuthor.aspx`}
            styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
          <DefaultButton text="Upload More" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({ wizardStep: 1, completedSteps: new Set() })} styles={{ root: { borderRadius: 4 } }} />
        </div>

        {/* Activity Log */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ padding: '12px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontWeight: 600, fontSize: 13, color: '#0f172a' }}>Activity Log</div>
          <div style={{ maxHeight: 300, overflowY: 'auto' }}>
            {activityLog.length === 0 ? (
              <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>No activity yet</div>
            ) : activityLog.map((entry, i) => (
              <div key={i} style={{ display: 'flex', gap: 10, padding: '8px 16px', borderBottom: '1px solid #f8fafc', fontSize: 12 }}>
                <span style={{ color: '#94a3b8', flexShrink: 0, minWidth: 60 }}>{entry.time.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                <span style={{ width: 6, height: 6, borderRadius: '50%', marginTop: 5, flexShrink: 0, background: entry.type === 'success' ? '#059669' : entry.type === 'error' ? '#dc2626' : entry.type === 'warning' ? '#d97706' : '#94a3b8' }} />
                <span style={{ color: entry.type === 'error' ? '#dc2626' : '#334155' }}>{entry.message}</span>
              </div>
            ))}
          </div>
        </div>
      </>
    );
  }

  // ============================================================================
  // BATCH PANEL
  // ============================================================================

  private renderBatchPanel(): React.ReactElement {
    const { showBatchPanel, batchTemplateId, batchCategory, batchRisk, selectedIds, fastTrackTemplates } = this.state;
    const templateOptions: IDropdownOption[] = [{ key: '', text: '— No template —' }, ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))];
    const hasTemplate = !!batchTemplateId;
    const hasMeta = !!batchCategory || !!batchRisk;

    return (
      <StyledPanel isOpen={showBatchPanel} onDismiss={() => this.setState({ showBatchPanel: false })}
        headerText={`Batch Assign (${selectedIds.size} selected)`} type={PanelType.smallFixedFar}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            <PrimaryButton text="Apply" disabled={!hasTemplate && !hasMeta}
              onClick={async () => {
                if (hasTemplate) { for (const id of selectedIds) await this.applyTemplateToItem(id, parseInt(batchTemplateId)); }
                this.setState({ showBatchPanel: false, batchTemplateId: '', batchCategory: '', batchRisk: '', selectedIds: new Set() });
              }}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showBatchPanel: false })} styles={{ root: { borderRadius: 4 } }} />
          </Stack>
        )} isFooterAtBottom={true}>
        <Stack tokens={{ childrenGap: 20 }} style={{ paddingTop: 16 }}>
          <Text style={{ fontSize: 13, color: '#64748b' }}>Apply a Fast Track Template to {selectedIds.size} selected polic{selectedIds.size !== 1 ? 'ies' : 'y'}.</Text>
          <div style={{ background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 4, padding: 16 }}>
            <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 13, display: 'block', marginBottom: 8 }}>Fast Track Template</Text>
            <Dropdown selectedKey={batchTemplateId} options={templateOptions} onChange={(_, opt) => this.setState({ batchTemplateId: String(opt?.key || '') })}
              placeholder="Select a template..." styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          </div>
          <Text style={{ fontSize: 12, color: '#94a3b8', textAlign: 'center' }}>— or set individual fields —</Text>
          <Dropdown label="Category" selectedKey={batchCategory} options={CATEGORY_OPTIONS} onChange={(_, opt) => this.setState({ batchCategory: String(opt?.key || '') })} styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          <Dropdown label="Compliance Risk" selectedKey={batchRisk} options={RISK_OPTIONS} onChange={(_, opt) => this.setState({ batchRisk: String(opt?.key || '') })} styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
        </Stack>
      </StyledPanel>
    );
  }
}
