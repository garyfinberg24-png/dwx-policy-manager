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
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';

// ============================================================================
// TYPES
// ============================================================================

type WizardStep = 1 | 2 | 3 | 4;
type ImportStatus = 'pending' | 'uploading' | 'uploaded' | 'classifying' | 'classified' | 'enriched' | 'failed';
type MatchConfidence = 'Strong' | 'Likely' | 'Possible' | 'None';

interface IFileMetadata { title?: string; author?: string; subject?: string; category?: string; keywords?: string; company?: string; description?: string; }
interface IFastTrackTemplate { Id: number; Title: string; ProfileName: string; PolicyCategory: string; ComplianceRisk: string; ReadTimeframe: string; RequiresAcknowledgement: boolean; RequiresQuiz: boolean; TargetDepartments: string; }

interface IBulkImportItem {
  id: string; spId?: number; fileName: string; fileSize: number; fileType: string; file?: File;
  documentUrl?: string; status: ImportStatus;
  existingMetadata?: IFileMetadata; hasExistingMetadata: boolean; useExistingMetadata: boolean;
  // Editable metadata (populated by AI, template, or manual entry)
  title: string; category: string; risk: string; department: string; summary: string;
  readTimeframe: string; requiresAck: boolean;
  // Template
  templateId?: number; templateName?: string; matchConfidence?: MatchConfidence;
  error?: string;
}

interface IActivityLogEntry { time: Date; message: string; type: 'info' | 'success' | 'warning' | 'error'; }

interface IPolicyBulkUploadState {
  loading: boolean; detectedRole: PolicyManagerRole | null;
  wizardStep: WizardStep; completedSteps: Set<number>;
  imports: IBulkImportItem[];
  uploading: boolean; uploadProgress: number;
  classifying: boolean; classifyProgress: number;
  selectedIds: Set<string>; searchQuery: string; filterType: string; groupBy: string;
  successMessage: string; errorMessage: string; dragOver: boolean;
  fastTrackTemplates: IFastTrackTemplate[]; templatesLoaded: boolean;
  activityLog: IActivityLogEntry[];
  showBatchPanel: boolean; batchTemplateId: string; batchCategory: string; batchRisk: string;
  importHistory: Array<{ date: string; fileCount: number; classified: number; templates: number }>;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const MAX_FILES = 50;
const MAX_FILE_SIZE = 25 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ['.docx', '.pdf', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.rtf', '.txt'];
const SESSION_KEY = 'pm_bulk_upload_state';

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' }, { key: 'IT Security', text: 'IT Security' }, { key: 'HR', text: 'Human Resources' },
  { key: 'Compliance', text: 'Compliance' }, { key: 'Data Protection', text: 'Data Protection' },
  { key: 'Health & Safety', text: 'Health & Safety' }, { key: 'Finance', text: 'Finance' },
  { key: 'Legal', text: 'Legal' }, { key: 'Operations', text: 'Operations' },
  { key: 'Governance', text: 'Governance' }, { key: 'Other', text: 'Other' }
];

const RISK_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' }, { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }
];

const WIZARD_STEPS = [
  { num: 1, title: 'Upload', desc: 'Add policy documents', icon: 'CloudUpload' },
  { num: 2, title: 'Review', desc: 'Check files & metadata', icon: 'ViewAll' },
  { num: 3, title: 'Enrich', desc: 'AI, templates, or manual', icon: 'Tag' },
  { num: 4, title: 'Finish', desc: 'Summary & next steps', icon: 'Accept' },
];

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyBulkUpload extends React.Component<IPolicyBulkUploadProps, IPolicyBulkUploadState> {
  private _isMounted = false;
  private _fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IPolicyBulkUploadProps) {
    super(props);
    this._fileInputRef = React.createRef();

    // Restore from sessionStorage if available
    let restored: Partial<IPolicyBulkUploadState> = {};
    try {
      const saved = sessionStorage.getItem(SESSION_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        restored = {
          imports: (parsed.imports || []).map((i: any) => ({ ...i, file: undefined })),
          wizardStep: parsed.wizardStep || 1,
          completedSteps: new Set(parsed.completedSteps || []),
          activityLog: (parsed.activityLog || []).map((e: any) => ({ ...e, time: new Date(e.time) })),
          importHistory: parsed.importHistory || [],
        };
      }
    } catch { /* no saved state */ }

    this.state = {
      loading: true, detectedRole: null,
      wizardStep: (restored.wizardStep as WizardStep) || 1,
      completedSteps: restored.completedSteps || new Set(),
      imports: restored.imports || [],
      uploading: false, uploadProgress: 0, classifying: false, classifyProgress: 0,
      selectedIds: new Set(), searchQuery: '', filterType: 'All', groupBy: 'None',
      successMessage: '', errorMessage: '', dragOver: false,
      fastTrackTemplates: [], templatesLoaded: false,
      activityLog: restored.activityLog || [],
      showBatchPanel: false, batchTemplateId: '', batchCategory: '', batchRisk: '',
      importHistory: restored.importHistory || [],
    };
  }

  public componentDidMount(): void { this._isMounted = true; this.detectRole(); }
  public componentWillUnmount(): void { this._isMounted = false; this.saveToSession(); }

  public componentDidUpdate(_: any, prevState: IPolicyBulkUploadState): void {
    // Auto-save to sessionStorage on meaningful state changes
    if (prevState.imports !== this.state.imports || prevState.wizardStep !== this.state.wizardStep || prevState.activityLog !== this.state.activityLog) {
      this.saveToSession();
    }
  }

  private saveToSession(): void {
    try {
      const { imports, wizardStep, completedSteps, activityLog, importHistory } = this.state;
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({
        imports: imports.map(i => ({ ...i, file: undefined })),
        wizardStep, completedSteps: Array.from(completedSteps),
        activityLog: activityLog.slice(0, 100),
        importHistory
      }));
    } catch { /* sessionStorage full or unavailable */ }
  }

  private async detectRole(): Promise<void> {
    try {
      const rs = new RoleDetectionService(this.props.sp);
      const ur = await rs.getCurrentUserRoles();
      const role = ur && ur.length > 0 ? getHighestPolicyRole(ur) : PolicyManagerRole.User;
      if (this._isMounted) this.setState({ detectedRole: role, loading: false });
    } catch { if (this._isMounted) this.setState({ detectedRole: PolicyManagerRole.Author, loading: false }); }
    this.loadFastTrackTemplates();
  }

  private log(message: string, type: 'info' | 'success' | 'warning' | 'error' = 'info'): void {
    this.setState(prev => ({ activityLog: [{ time: new Date(), message, type }, ...prev.activityLog.slice(0, 199)] }));
  }

  private updateItem(id: string, updates: Partial<IBulkImportItem>): void {
    this.setState({ imports: this.state.imports.map(i => i.id === id ? { ...i, ...updates } : i) });
  }

  // ============================================================================
  // FILE METADATA EXTRACTION
  // ============================================================================

  private async extractFileMetadata(file: File): Promise<IFileMetadata> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    const meta: IFileMetadata = {};
    try {
      if (['docx', 'pptx', 'xlsx'].includes(ext)) {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let str = ''; for (let i = 0; i < Math.min(bytes.length, 256000); i++) str += String.fromCharCode(bytes[i]);
        const get = (tag: string): string => { for (const p of [`dc:${tag}`, `cp:${tag}`, `dcterms:${tag}`]) { const m = str.match(new RegExp(`<${p}[^>]*>([^<]+)</${p}>`, 'i')); if (m) return m[1].trim(); } return ''; };
        meta.title = get('title'); meta.author = get('creator') || get('lastModifiedBy');
        meta.subject = get('subject'); meta.category = get('category'); meta.keywords = get('keywords');
        const cm = str.match(/<Company>([^<]+)<\/Company>/i); if (cm) meta.company = cm[1].trim();
      }
      if (ext === 'pdf') {
        const t = await file.slice(0, 4096).text();
        const g = (k: string) => { const m = t.match(new RegExp(`/${k}\\s*\\(([^)]+)\\)`, 'i')); return m ? m[1].trim() : ''; };
        meta.title = g('Title'); meta.author = g('Author'); meta.subject = g('Subject'); meta.keywords = g('Keywords');
      }
    } catch { /* extraction failed */ }
    return meta;
  }

  private async extractTextFromFile(file: File): Promise<string> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    try {
      if (['txt', 'rtf', 'csv', 'md'].includes(ext)) return (await file.text()).substring(0, 5000);
      if (['docx', 'doc', 'pptx', 'xlsx'].includes(ext)) {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let str = ''; for (let i = 0; i < Math.min(bytes.length, 512000); i++) str += String.fromCharCode(bytes[i]);
        const parts: string[] = []; const re = /<(?:w:|a:)?t[^>]*>([^<]+)<\/(?:w:|a:)?t>/g; let m: RegExpExecArray | null;
        while ((m = re.exec(str)) !== null && parts.length < 800) { if (m[1].trim()) parts.push(m[1].trim()); }
        if (parts.length > 5) return parts.join(' ').substring(0, 5000);
      }
      if (ext === 'pdf') {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let t = ''; for (let i = 0; i < Math.min(bytes.length, 512000); i++) { const c = bytes[i]; t += (c >= 32 && c <= 126) || c === 10 || c === 13 ? String.fromCharCode(c) : ' '; }
        const s = t.replace(/\s{3,}/g, ' ').split(/\s{2,}/).filter(s => s.length > 20);
        if (s.length > 3) return s.join(' ').substring(0, 5000);
      }
    } catch { /* */ }
    return `Document: ${file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')}`;
  }

  // ============================================================================
  // STEP 1: FILE HANDLING + UPLOAD
  // ============================================================================

  private handleFileDrop = (e: React.DragEvent): void => { e.preventDefault(); e.stopPropagation(); this.setState({ dragOver: false }); this.processFiles(Array.from(e.dataTransfer.files)); };
  private handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => { this.processFiles(Array.from(e.target.files || [])); if (this._fileInputRef.current) this._fileInputRef.current.value = ''; };

  private async processFiles(files: File[]): Promise<void> {
    const { imports } = this.state;
    const remaining = MAX_FILES - imports.length;
    if (remaining <= 0) { this.setState({ errorMessage: `Maximum ${MAX_FILES} files.` }); return; }
    const newItems: IBulkImportItem[] = [];
    for (const file of files.slice(0, remaining)) {
      const ext = '.' + file.name.split('.').pop()?.toLowerCase();
      if (!ALLOWED_EXTENSIONS.includes(ext)) { this.log(`${file.name}: unsupported`, 'warning'); continue; }
      if (file.size > MAX_FILE_SIZE) { this.log(`${file.name}: too large`, 'warning'); continue; }
      if (imports.some(i => i.fileName === file.name)) { this.log(`${file.name}: duplicate`, 'warning'); continue; }
      const meta = await this.extractFileMetadata(file);
      const hasMeta = !!(meta.title || meta.category || meta.keywords);
      newItems.push({
        id: `imp_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        fileName: file.name, fileSize: file.size, fileType: ext, file, status: 'pending',
        existingMetadata: meta, hasExistingMetadata: hasMeta, useExistingMetadata: false,
        title: meta.title || file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' '),
        category: meta.category || '', risk: '', department: '', summary: '',
        readTimeframe: '', requiresAck: false,
      });
      this.log(`Added: ${file.name}${hasMeta ? ' (metadata found)' : ''}`, hasMeta ? 'success' : 'info');
    }
    this.setState({ imports: [...imports, ...newItems] });
  }

  private removeImport = (id: string): void => {
    const item = this.state.imports.find(i => i.id === id);
    if (item) this.log(`Removed: ${item.fileName}`, 'info');
    this.setState(prev => ({ imports: prev.imports.filter(i => i.id !== id), selectedIds: (() => { const s = new Set(prev.selectedIds); s.delete(id); return s; })() }));
  };

  private async uploadToSharePoint(): Promise<void> {
    const { imports } = this.state;
    const toUpload = imports.filter(i => i.status === 'pending' && i.file);
    if (toUpload.length === 0) return;
    this.setState({ uploading: true, uploadProgress: 0 });
    this.log(`Uploading ${toUpload.length} files...`, 'info');
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
    const siteRelUrl = this.props.context?.pageContext?.web?.serverRelativeUrl || '/sites/PolicyManager';

    let digest = '';
    try {
      const { SPHttpClient } = await import('@microsoft/sp-http');
      const r = await this.props.context.spHttpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {});
      const d = await r.json(); digest = d?.FormDigestValue || d?.d?.GetContextWebInformation?.FormDigestValue || '';
    } catch { digest = (document.getElementById('__REQUESTDIGEST') as HTMLInputElement)?.value || ''; }

    const folderUrl = `${siteRelUrl}/${PM_LISTS.POLICY_SOURCE_DOCUMENTS}/BulkImports`;
    try { const x = new XMLHttpRequest(); x.open('POST', `${siteUrl}/_api/web/folders`, false); x.setRequestHeader('Accept', 'application/json; odata=verbose'); x.setRequestHeader('Content-Type', 'application/json; odata=verbose'); x.setRequestHeader('X-RequestDigest', digest); x.send(JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': folderUrl })); } catch { /* exists */ }

    let processed = 0, succeeded = 0;
    for (const item of toUpload) {
      try {
        this.updateItem(item.id, { status: 'uploading' });
        const buf = await item.file.arrayBuffer();
        const safeName = item.fileName.replace(/[#%&*:<>?\/\\{|}~]/g, '_');
        const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folderUrl}')/Files/AddUsingPath(decodedurl='${encodeURIComponent(safeName)}',overwrite=true)`;
        const docUrl: string = await new Promise((res, rej) => {
          const x = new XMLHttpRequest(); x.open('POST', endpoint, true);
          x.setRequestHeader('Accept', 'application/json; odata=verbose'); x.setRequestHeader('Content-Type', 'application/octet-stream'); x.setRequestHeader('X-RequestDigest', digest);
          x.responseType = 'json'; x.onload = () => x.status >= 200 && x.status < 300 ? res(x.response?.d?.ServerRelativeUrl || '') : rej(new Error(`${x.status}`)); x.onerror = () => rej(new Error('Network')); x.send(new Uint8Array(buf));
        });
        const title = item.title || item.fileName.replace(/\.[^.]+$/, '');
        const result = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.add({ Title: title, PolicyStatus: 'Draft' });
        const spId = result?.data?.Id || result?.data?.id;
        if (spId && docUrl) { try { await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(spId).update({ PolicyName: title, DocumentURL: docUrl }); } catch { /* */ } }
        this.updateItem(item.id, { spId, documentUrl: docUrl, status: 'uploaded' });
        this.log(`Uploaded: ${item.fileName}`, 'success'); succeeded++;
      } catch (err) {
        this.updateItem(item.id, { status: 'failed', error: String(err) });
        this.log(`Failed: ${item.fileName}`, 'error');
      }
      processed++;
      if (this._isMounted) this.setState({ uploadProgress: Math.round((processed / toUpload.length) * 100) });
    }
    const completed = new Set(this.state.completedSteps); completed.add(1);
    this.setState({ uploading: false, completedSteps: completed, wizardStep: 2 });
    this.log(`Upload complete: ${succeeded}/${toUpload.length}`, succeeded === toUpload.length ? 'success' : 'warning');
  }

  // ============================================================================
  // STEP 3: AI CLASSIFICATION
  // ============================================================================

  private async classifySelected(): Promise<void> {
    const { imports, selectedIds } = this.state;
    const toClassify = imports.filter(i => selectedIds.has(i.id) && ['uploaded', 'classified', 'enriched'].includes(i.status));
    if (toClassify.length === 0) { this.setState({ errorMessage: 'Select files to classify.' }); return; }
    this.setState({ classifying: true, classifyProgress: 0 });
    this.log(`Classifying ${toClassify.length} files with AI...`, 'info');

    let functionUrl = '';
    try { const c = await this.props.sp.web.lists.getByTitle('PM_Configuration').items.filter("ConfigKey eq 'Integration.AI.Chat.FunctionUrl'").select('ConfigValue').top(1)(); functionUrl = c[0]?.ConfigValue || ''; } catch { /* */ }
    if (!functionUrl) functionUrl = localStorage.getItem('PM_AI_ChatFunctionUrl') || '';

    let processed = 0;
    for (const item of toClassify) {
      try {
        this.updateItem(item.id, { status: 'classifying' });
        let content = ''; if (item.file) { try { content = await this.extractTextFromFile(item.file); } catch { /* */ } }
        let suggestions: any = {};
        if (functionUrl) {
          const ctx = content.length > 100 ? `\n\nDocument content:\n"""${content.substring(0, 3000)}"""` : '';
          try {
            const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), 45000);
            const resp = await fetch(functionUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ mode: 'author-assist', message: `Classify this policy document. Filename: "${item.fileName}"${ctx}\n\nExtract:\n1. Title (clean professional title)\n2. PolicyCategory (IT Security, HR, Compliance, Data Protection, Health & Safety, Finance, Legal, Operations, Governance, Other)\n3. ComplianceRisk (Critical, High, Medium, Low, Informational)\n4. Departments (comma-separated)\n5. Summary (1-2 sentences)\n6. ReadTimeframe (Immediate, Day 1, Day 3, Week 1, Week 2, Month 1)\n\nRespond ONLY with JSON: {title, category, risk, departments, summary, readTimeframe}`, history: [], context: [] }),
              signal: ctrl.signal });
            clearTimeout(t);
            if (resp.ok) { const d = await resp.json(); const c = d?.response || d?.content || ''; try { const m = c.match(/\{[\s\S]*\}/); if (m) suggestions = JSON.parse(m[0]); } catch { /* */ } }
          } catch { /* AI failed */ }
        }
        if (!suggestions.category) suggestions = this.heuristicClassify(item.fileName, item.title);

        this.updateItem(item.id, {
          status: 'classified',
          title: suggestions.title || item.title,
          category: suggestions.category || item.category || 'Other',
          risk: suggestions.risk || item.risk || 'Medium',
          department: Array.isArray(suggestions.departments) ? suggestions.departments.join(', ') : (suggestions.departments || item.department || 'All Employees'),
          summary: suggestions.summary || item.summary || '',
          readTimeframe: suggestions.readTimeframe || item.readTimeframe || 'Week 1',
          file: undefined, // free memory
        });

        // Match template
        const match = this.matchTemplate(suggestions.category || 'Other', suggestions.risk || 'Medium');
        if (match) this.updateItem(item.id, { templateId: match.templateId, templateName: match.templateName, matchConfidence: match.confidence });

        this.log(`Classified: ${item.fileName} → ${suggestions.category || 'Other'}`, 'success');
      } catch (err) {
        this.updateItem(item.id, { status: 'uploaded', error: String(err) });
        this.log(`Classification failed: ${item.fileName}`, 'error');
      }
      processed++;
      if (this._isMounted) this.setState({ classifyProgress: Math.round((processed / toClassify.length) * 100) });
    }
    this.setState({ classifying: false }); this.log(`Classification complete`, 'success');
  }

  private heuristicClassify(fileName: string, title: string): any {
    const t = (fileName + ' ' + title).toLowerCase();
    let cat = 'Other', risk = 'Medium';
    if (/security|cyber|access|password|firewall/i.test(t)) { cat = 'IT Security'; risk = 'High'; }
    else if (/hr|human|employee|leave|conduct/i.test(t)) { cat = 'HR'; risk = 'Medium'; }
    else if (/compliance|regulat|audit/i.test(t)) { cat = 'Compliance'; risk = 'High'; }
    else if (/data|privacy|gdpr|pii/i.test(t)) { cat = 'Data Protection'; risk = 'Critical'; }
    else if (/health|safety|incident/i.test(t)) { cat = 'Health & Safety'; risk = 'High'; }
    else if (/financ|expense|budget|travel/i.test(t)) { cat = 'Finance'; risk = 'Medium'; }
    else if (/legal|contract|nda/i.test(t)) { cat = 'Legal'; risk = 'High'; }
    else if (/operat|process|procedure/i.test(t)) { cat = 'Operations'; risk = 'Low'; }
    else if (/govern|board|ethic/i.test(t)) { cat = 'Governance'; risk = 'High'; }
    return { title: title || fileName.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ').replace(/\b\w/g, c => c.toUpperCase()), category: cat, risk, departments: 'All Employees', summary: '', readTimeframe: 'Week 1' };
  }

  // ============================================================================
  // TEMPLATES
  // ============================================================================

  private async loadFastTrackTemplates(): Promise<void> {
    if (this.state.templatesLoaded) return;
    try {
      let items: any[] = [];
      try { items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES).items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk', 'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz', 'TargetDepartments', 'IsActive').orderBy('Title').top(100)(); }
      catch { try { items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES).items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk').orderBy('Title').top(100)(); } catch { /* */ } }
      const templates = items.filter((t: any) => t.IsActive !== false).map((t: any) => ({ Id: t.Id, Title: t.Title || t.ProfileName || `Template ${t.Id}`, ProfileName: t.ProfileName || '', PolicyCategory: t.PolicyCategory || '', ComplianceRisk: t.ComplianceRisk || 'Medium', ReadTimeframe: t.ReadTimeframe || 'Week 1', RequiresAcknowledgement: t.RequiresAcknowledgement !== false, RequiresQuiz: t.RequiresQuiz || false, TargetDepartments: t.TargetDepartments || '' }));
      if (this._isMounted) this.setState({ fastTrackTemplates: templates, templatesLoaded: true });
    } catch { if (this._isMounted) this.setState({ templatesLoaded: true }); }
  }

  private matchTemplate(category: string, risk: string): { templateId: number; templateName: string; confidence: MatchConfidence } | null {
    const { fastTrackTemplates } = this.state;
    if (fastTrackTemplates.length === 0) return null;
    let best: IFastTrackTemplate | null = null, bestScore = 0;
    for (const t of fastTrackTemplates) {
      let s = 0; if (t.PolicyCategory.toLowerCase() === category.toLowerCase()) s += 3; if (t.ComplianceRisk.toLowerCase() === risk.toLowerCase()) s += 2;
      if (s > bestScore) { bestScore = s; best = t; }
    }
    if (!best || bestScore === 0) return null;
    return { templateId: best.Id, templateName: best.Title, confidence: bestScore >= 5 ? 'Strong' : bestScore >= 3 ? 'Likely' : 'Possible' };
  }

  private applyTemplate(itemId: string, templateId: number): void {
    const template = this.state.fastTrackTemplates.find(t => t.Id === templateId);
    if (!template) return;
    this.updateItem(itemId, {
      category: template.PolicyCategory, risk: template.ComplianceRisk, readTimeframe: template.ReadTimeframe,
      requiresAck: template.RequiresAcknowledgement, department: template.TargetDepartments || 'All Employees',
      templateId: template.Id, templateName: template.Title, status: 'enriched',
    });
    // Write to SP if spId exists
    const item = this.state.imports.find(i => i.id === itemId);
    if (item?.spId) {
      this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({
        PolicyCategory: template.PolicyCategory, ComplianceRisk: template.ComplianceRisk
      }).catch(() => { /* best effort */ });
    }
    this.log(`Template: ${template.Title} → ${item?.fileName || itemId}`, 'success');
  }

  // ============================================================================
  // RENDER: MAIN
  // ============================================================================

  public render(): React.ReactElement {
    const { detectedRole, loading } = this.state;
    if (loading) return (<ErrorBoundary fallbackMessage="Error"><JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}><div style={{ padding: 60, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading..." /></div></JmlAppLayout></ErrorBoundary>);
    if (detectedRole !== null && !hasMinimumRole(detectedRole, PolicyManagerRole.Author)) return (<ErrorBoundary fallbackMessage="Error"><JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}><section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}><Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} /><Text variant="xLarge" block styles={{ root: { fontWeight: 600 } }}>Access Denied</Text></section></JmlAppLayout></ErrorBoundary>);

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Bulk Upload.">
        <JmlAppLayout title={this.props.title || 'Policy Bulk Upload'} context={this.props.context} sp={this.props.sp} activeNavKey="bulk-upload" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}>
          {this.renderWizard()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ============================================================================
  // RENDER: WIZARD
  // ============================================================================

  private renderWizard(): React.ReactElement {
    const { wizardStep, completedSteps, imports, successMessage, errorMessage } = this.state;
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    return (
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', minHeight: 'calc(100vh - 180px)', background: '#fff', borderRadius: 10, overflow: 'hidden', border: '1px solid #e2e8f0', margin: '0 auto', maxWidth: 1400 }}>
        {/* Sidebar */}
        <aside style={{ background: '#fff', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column' }}>
          <div style={{ padding: '20px 16px 14px', borderBottom: '1px solid #e2e8f0' }}>
            <Text style={{ fontSize: 15, fontWeight: 700, color: '#0f172a', display: 'block' }}>Bulk Upload</Text>
            <Text style={{ fontSize: 10, color: '#94a3b8', marginTop: 2, display: 'block' }}>4 steps to import policies</Text>
          </div>
          <div style={{ flex: 1, padding: '6px 0' }}>
            {WIZARD_STEPS.map(step => {
              const done = completedSteps.has(step.num); const active = step.num === wizardStep; const clickable = done || step.num <= wizardStep;
              return (
                <div key={step.num} onClick={() => clickable && this.setState({ wizardStep: step.num as WizardStep })}
                  style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 16px', cursor: clickable ? 'pointer' : 'default', borderLeft: active ? '3px solid #0d9488' : '3px solid transparent', background: active ? '#f0fdfa' : 'transparent' }}>
                  <div style={{ width: 26, height: 26, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, flexShrink: 0, background: done ? '#0d9488' : active ? '#f0fdfa' : '#fff', color: done ? '#fff' : active ? '#0d9488' : '#94a3b8', border: `2px solid ${done ? '#0d9488' : active ? '#0d9488' : '#e2e8f0'}` }}>
                    {done ? <Icon iconName="CheckMark" style={{ fontSize: 10 }} /> : step.num}
                  </div>
                  <div><div style={{ fontWeight: active ? 600 : 500, color: active ? '#0d9488' : done ? '#0f172a' : '#475569', fontSize: 12 }}>{step.title}</div><div style={{ fontSize: 9, color: '#94a3b8' }}>{step.desc}</div></div>
                </div>
              );
            })}
          </div>
          <div style={{ padding: '12px 16px', borderTop: '1px solid #e2e8f0', fontSize: 11 }}>
            {[{ l: 'Total', v: imports.length, c: '#475569' }, { l: 'Uploaded', v: imports.filter(i => !['pending', 'failed'].includes(i.status)).length, c: '#2563eb' }, { l: 'Enriched', v: imports.filter(i => ['classified', 'enriched'].includes(i.status)).length, c: '#7c3aed' }, { l: 'With Template', v: imports.filter(i => !!i.templateId).length, c: '#059669' }].map(k =>
              <div key={k.l} style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0' }}><span style={{ color: '#64748b' }}>{k.l}</span><span style={{ fontWeight: 700, color: k.c }}>{k.v}</span></div>
            )}
          </div>
        </aside>

        {/* Content */}
        <div style={{ padding: '24px 32px', overflowY: 'auto', background: '#f8fafc' }}>
          {successMessage && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ successMessage: '' })} style={{ marginBottom: 12 }}>{successMessage}</MessageBar>}
          {errorMessage && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ errorMessage: '' })} style={{ marginBottom: 12 }}>{errorMessage}</MessageBar>}

          {wizardStep === 1 && this.renderStep1_Upload()}
          {wizardStep === 2 && this.renderStep2_Review()}
          {wizardStep === 3 && this.renderStep3_Enrich(siteUrl)}
          {wizardStep === 4 && this.renderStep4_Finish(siteUrl)}

          {/* Nav footer */}
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 20, paddingTop: 14, borderTop: '1px solid #e2e8f0' }}>
            <DefaultButton text="Previous" iconProps={{ iconName: 'ChevronLeft' }} disabled={wizardStep === 1}
              onClick={() => this.setState({ wizardStep: Math.max(1, wizardStep - 1) as WizardStep })}
              styles={{ root: { borderRadius: 4, visibility: wizardStep === 1 ? 'hidden' : 'visible' } }} />
            {wizardStep < 4 && (
              <PrimaryButton onClick={() => { const c = new Set(this.state.completedSteps); c.add(wizardStep); this.setState({ wizardStep: Math.min(4, wizardStep + 1) as WizardStep, completedSteps: c }); }}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}>
                Next <Icon iconName="ChevronRight" style={{ marginLeft: 6 }} />
              </PrimaryButton>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // STEP 1: UPLOAD
  // ============================================================================

  private renderStep1_Upload(): React.ReactElement {
    const { imports, uploading, uploadProgress, dragOver } = this.state;
    const pending = imports.filter(i => i.status === 'pending');
    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Upload Policy Documents</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 16px' }}>Drag and drop or browse. DOCX, PDF, XLSX, PPTX — up to 25MB, max {MAX_FILES} files.</p>

        <div onDragOver={(e) => { e.preventDefault(); this.setState({ dragOver: true }); }} onDragLeave={() => this.setState({ dragOver: false })} onDrop={this.handleFileDrop}
          onClick={() => this._fileInputRef.current?.click()}
          style={{ border: `2px dashed ${dragOver ? '#0d9488' : '#cbd5e1'}`, borderRadius: 10, padding: '36px 24px', textAlign: 'center', cursor: 'pointer', background: dragOver ? '#f0fdfa' : '#fff', marginBottom: 16 }}>
          <input ref={this._fileInputRef} type="file" multiple accept={ALLOWED_EXTENSIONS.join(',')} onChange={this.handleFileSelect} style={{ display: 'none' }} />
          <svg viewBox="0 0 24 24" fill="none" width="32" height="32" style={{ color: dragOver ? '#0d9488' : '#94a3b8', marginBottom: 8 }}><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /><path d="M17 8l-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /><path d="M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /></svg>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{dragOver ? 'Drop files here' : 'Drag & drop or click to browse'}</div>
        </div>

        {imports.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 16 }}>
            {imports.map(item => {
              const sz = item.fileSize < 1048576 ? `${Math.round(item.fileSize / 1024)} KB` : `${(item.fileSize / 1048576).toFixed(1)} MB`;
              return (
                <div key={item.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 16px', borderBottom: '1px solid #f1f5f9', fontSize: 13 }}>
                  <div style={{ flex: 1 }}><span style={{ fontWeight: 600, color: '#0f172a' }}>{item.title}</span><span style={{ color: '#94a3b8', marginLeft: 8, fontSize: 11 }}>{item.fileName} · {sz}</span></div>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: item.hasExistingMetadata ? '#f0fdf4' : '#f1f5f9', color: item.hasExistingMetadata ? '#059669' : '#94a3b8' }}>{item.hasExistingMetadata ? 'Meta' : 'None'}</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: item.status === 'uploaded' ? '#f0fdf4' : item.status === 'failed' ? '#fef2f2' : '#f1f5f9', color: item.status === 'uploaded' ? '#059669' : item.status === 'failed' ? '#dc2626' : '#94a3b8' }}>{item.status === 'uploaded' ? '✓' : item.status === 'failed' ? '✕' : '○'}</span>
                  <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => this.removeImport(item.id)} styles={{ root: { width: 22, height: 22 }, icon: { fontSize: 10, color: '#dc2626' } }} />
                </div>
              );
            })}
          </div>
        )}

        {uploading && <ProgressIndicator label={`Uploading... ${uploadProgress}%`} percentComplete={uploadProgress / 100} style={{ marginBottom: 12 }} />}
        {pending.length > 0 && !uploading && (
          <PrimaryButton text={`Upload ${pending.length} File${pending.length !== 1 ? 's' : ''}`} iconProps={{ iconName: 'CloudUpload' }} onClick={() => this.uploadToSharePoint()}
            styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
        )}
      </>
    );
  }

  // ============================================================================
  // STEP 2: REVIEW
  // ============================================================================

  private renderStep2_Review(): React.ReactElement {
    const { imports, searchQuery, selectedIds } = this.state;
    const uploaded = imports.filter(i => !['pending', 'failed'].includes(i.status));
    let filtered = uploaded;
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); filtered = filtered.filter(i => i.title.toLowerCase().includes(q) || i.fileName.toLowerCase().includes(q)); }
    const allSelected = filtered.length > 0 && filtered.every(i => selectedIds.has(i.id));

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Review Uploaded Documents</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 14px' }}>Check existing metadata, edit titles. Select files for the next step.</p>

        <div style={{ display: 'flex', gap: 10, marginBottom: 12, alignItems: 'center' }}>
          <SearchBox placeholder="Search..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 220 } }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: '#64748b' }}>{selectedIds.size} selected · {filtered.length} files</span>
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '32px 1fr 90px 90px 90px 80px', padding: '8px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" checked={allSelected} onChange={() => { if (allSelected) this.setState({ selectedIds: new Set() }); else this.setState({ selectedIds: new Set(filtered.map(i => i.id)) }); }} /></div>
            <div>Title / File</div><div>Author</div><div>Category</div><div>Keywords</div><div>Metadata</div>
          </div>
          {filtered.map(item => {
            const meta = item.existingMetadata || {};
            return (
              <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '32px 1fr 90px 90px 90px 80px', padding: '8px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: selectedIds.has(item.id) ? '#f0fdfa' : '#fff' }}>
                <div><input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => { const n = new Set(selectedIds); if (n.has(item.id)) n.delete(item.id); else n.add(item.id); this.setState({ selectedIds: n }); }} /></div>
                <div>
                  <input type="text" value={item.title} onChange={(e) => this.updateItem(item.id, { title: (e.target as HTMLInputElement).value })}
                    style={{ width: '100%', border: 'none', background: 'transparent', fontSize: 13, fontWeight: 600, color: '#0f172a', outline: 'none', padding: '2px 0' }} />
                  <div style={{ fontSize: 10, color: '#94a3b8' }}>{item.fileName} · {item.fileType.replace('.', '').toUpperCase()}</div>
                </div>
                <div style={{ fontSize: 11, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{meta.author || '—'}</div>
                <div style={{ fontSize: 11, color: '#475569' }}>{meta.category || '—'}</div>
                <div style={{ fontSize: 11, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{meta.keywords || '—'}</div>
                <div>{item.hasExistingMetadata ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Found</span> : <span style={{ fontSize: 10, color: '#cbd5e1' }}>—</span>}</div>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // ============================================================================
  // STEP 3: ENRICH METADATA (merged Classify + Templates)
  // ============================================================================

  private renderStep3_Enrich(siteUrl: string): React.ReactElement {
    const { imports, classifying, classifyProgress, selectedIds, fastTrackTemplates } = this.state;
    const enrichable = imports.filter(i => !['pending', 'failed'].includes(i.status));
    const templateOptions: IDropdownOption[] = [{ key: '', text: '— No template —' }, ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))];
    const selectedCount = selectedIds.size;
    const riskColor = (r: string) => r === 'Critical' ? '#dc2626' : r === 'High' ? '#d97706' : r === 'Medium' ? '#0d9488' : r === 'Low' ? '#059669' : '#94a3b8';

    return (
      <>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 14 }}>
          <div>
            <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Enrich Metadata</h2>
            <p style={{ fontSize: 13, color: '#64748b', margin: 0 }}>Edit fields directly, use AI to auto-fill, or apply a Fast Track Template. All optional.</p>
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            {selectedCount > 0 && !classifying && (
              <>
                <PrimaryButton text={`AI Classify (${selectedCount})`} iconProps={{ iconName: 'Processing' }}
                  onClick={() => this.classifySelected()}
                  styles={{ root: { background: '#7c3aed', borderColor: '#7c3aed', borderRadius: 4, fontSize: 12, height: 30 }, rootHovered: { background: '#6d28d9', borderColor: '#6d28d9' } }} />
                <DefaultButton text={`Batch Template (${selectedCount})`} iconProps={{ iconName: 'Tag' }}
                  onClick={() => this.setState({ showBatchPanel: true })}
                  styles={{ root: { borderRadius: 4, fontSize: 12, height: 30 } }} />
              </>
            )}
          </div>
        </div>

        {classifying && <ProgressIndicator label={`Classifying... ${classifyProgress}%`} percentComplete={classifyProgress / 100} style={{ marginBottom: 12 }} />}

        {/* Editable enrichment table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '32px 1fr 120px 80px 100px 170px 70px 50px', padding: '8px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" onChange={(e) => { if ((e.target as HTMLInputElement).checked) this.setState({ selectedIds: new Set(enrichable.map(i => i.id)) }); else this.setState({ selectedIds: new Set() }); }} /></div>
            <div>Policy Title</div><div>Category</div><div>Risk</div><div>Department</div><div>Fast Track Template</div><div>Status</div><div></div>
          </div>
          {enrichable.map(item => {
            const isClassifying = item.status === 'classifying';
            return (
              <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '32px 1fr 120px 80px 100px 170px 70px 50px', padding: '8px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: isClassifying ? '#faf5ff' : selectedIds.has(item.id) ? '#f0fdfa' : '#fff', opacity: isClassifying ? 0.7 : 1 }}>
                <div><input type="checkbox" checked={selectedIds.has(item.id)} disabled={isClassifying} onChange={() => { const n = new Set(selectedIds); if (n.has(item.id)) n.delete(item.id); else n.add(item.id); this.setState({ selectedIds: n }); }} /></div>

                {/* Title — editable */}
                <div>
                  <input type="text" value={item.title} onChange={(e) => this.updateItem(item.id, { title: (e.target as HTMLInputElement).value })} disabled={isClassifying}
                    style={{ width: '100%', border: 'none', background: 'transparent', fontSize: 13, fontWeight: 600, color: '#0f172a', outline: 'none', padding: '2px 0' }} />
                  <div style={{ fontSize: 10, color: '#cbd5e1' }}>{item.fileName}</div>
                </div>

                {/* Category — editable dropdown */}
                <div>
                  <Dropdown selectedKey={item.category} options={CATEGORY_OPTIONS} disabled={isClassifying}
                    onChange={(_, opt) => this.updateItem(item.id, { category: String(opt?.key || ''), status: item.status === 'uploaded' ? 'enriched' : item.status })}
                    styles={{ root: { minWidth: 0 }, title: { fontSize: 11, height: 26, lineHeight: '24px', borderRadius: 4, borderColor: '#e2e8f0' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }} />
                </div>

                {/* Risk — editable dropdown */}
                <div>
                  <Dropdown selectedKey={item.risk} options={RISK_OPTIONS} disabled={isClassifying}
                    onChange={(_, opt) => this.updateItem(item.id, { risk: String(opt?.key || ''), status: item.status === 'uploaded' ? 'enriched' : item.status })}
                    styles={{ root: { minWidth: 0 }, title: { fontSize: 11, height: 26, lineHeight: '24px', borderRadius: 4, borderColor: '#e2e8f0', color: riskColor(item.risk) }, caretDownWrapper: { height: 26, lineHeight: '26px' } }} />
                </div>

                {/* Department — editable text */}
                <div>
                  <input type="text" value={item.department} placeholder="All" disabled={isClassifying}
                    onChange={(e) => this.updateItem(item.id, { department: (e.target as HTMLInputElement).value })}
                    style={{ width: '100%', border: 'none', background: 'transparent', fontSize: 11, color: '#475569', outline: 'none', padding: '2px 0' }} />
                </div>

                {/* Fast Track Template — dropdown */}
                <div>
                  <Dropdown selectedKey={item.templateId ? String(item.templateId) : ''} options={templateOptions} disabled={isClassifying}
                    onChange={(_, opt) => { if (opt?.key) this.applyTemplate(item.id, parseInt(String(opt.key))); else this.updateItem(item.id, { templateId: undefined, templateName: undefined }); }}
                    styles={{ root: { minWidth: 0 }, title: { fontSize: 11, height: 26, lineHeight: '24px', borderRadius: 4, borderColor: item.templateId ? '#bbf7d0' : '#e2e8f0', background: item.templateId ? '#f0fdf4' : '#fff' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }} />
                </div>

                {/* Status */}
                <div>
                  {isClassifying ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f5f3ff', color: '#7c3aed' }}>AI...</span> :
                    item.status === 'classified' || item.status === 'enriched' ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Done</span> :
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f1f5f9', color: '#94a3b8' }}>Draft</span>}
                </div>

                {/* Open in builder */}
                <div>
                  {item.spId && <IconButton iconProps={{ iconName: 'Edit' }} title="Open in Policy Builder" href={`${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${item.spId}`}
                    styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 12, color: '#0d9488' } }} />}
                </div>
              </div>
            );
          })}
          {enrichable.length === 0 && <div style={{ padding: 32, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No uploaded files. Go back to Upload.</div>}
        </div>

        {this.renderBatchPanel()}
      </>
    );
  }

  // ============================================================================
  // STEP 4: FINISH
  // ============================================================================

  private renderStep4_Finish(siteUrl: string): React.ReactElement {
    const { imports, activityLog, importHistory } = this.state;
    const uploaded = imports.filter(i => !['pending', 'failed'].includes(i.status)).length;
    const enriched = imports.filter(i => ['classified', 'enriched'].includes(i.status)).length;
    const templated = imports.filter(i => !!i.templateId).length;
    const failed = imports.filter(i => i.status === 'failed').length;

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Import Summary</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 20px' }}>Your policies are imported as drafts. Open them in the Policy Builder to add content and submit.</p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 20 }}>
          {[{ l: 'Uploaded', v: uploaded, c: '#2563eb' }, { l: 'Enriched', v: enriched, c: '#7c3aed' }, { l: 'With Template', v: templated, c: '#059669' }, { l: 'Failed', v: failed, c: failed > 0 ? '#dc2626' : '#94a3b8' }].map(k =>
            <div key={k.l} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.c}`, padding: '14px 16px', textAlign: 'center' }}>
              <div style={{ fontSize: 24, fontWeight: 700, color: k.c }}>{k.v}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.l}</div>
            </div>
          )}
        </div>

        <div style={{ display: 'flex', gap: 8, marginBottom: 20 }}>
          <PrimaryButton text="Open Drafts & Pipeline" iconProps={{ iconName: 'ViewAll' }} href={`${siteUrl}/SitePages/PolicyAuthor.aspx`}
            styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
          <DefaultButton text="New Import" iconProps={{ iconName: 'Add' }}
            onClick={() => {
              // Save current batch to history
              const batch = { date: new Date().toISOString(), fileCount: imports.length, classified: enriched, templates: templated };
              this.setState({ wizardStep: 1, completedSteps: new Set(), imports: [], selectedIds: new Set(), activityLog: [], importHistory: [batch, ...this.state.importHistory.slice(0, 19)] });
              sessionStorage.removeItem(SESSION_KEY);
            }}
            styles={{ root: { borderRadius: 4 } }} />
        </div>

        {/* Activity Log */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 }}>
          <div style={{ padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontWeight: 600, fontSize: 13, color: '#0f172a' }}>Activity Log</div>
          <div style={{ maxHeight: 250, overflowY: 'auto' }}>
            {activityLog.length === 0 ? <div style={{ padding: 20, textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>No activity</div> :
              activityLog.map((e, i) => (
                <div key={i} style={{ display: 'flex', gap: 8, padding: '6px 16px', borderBottom: '1px solid #f8fafc', fontSize: 11 }}>
                  <span style={{ color: '#94a3b8', flexShrink: 0, minWidth: 55 }}>{e.time.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                  <span style={{ width: 6, height: 6, borderRadius: '50%', marginTop: 4, flexShrink: 0, background: e.type === 'success' ? '#059669' : e.type === 'error' ? '#dc2626' : e.type === 'warning' ? '#d97706' : '#94a3b8' }} />
                  <span style={{ color: e.type === 'error' ? '#dc2626' : '#334155' }}>{e.message}</span>
                </div>
              ))}
          </div>
        </div>

        {/* Import History */}
        {importHistory.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{ padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontWeight: 600, fontSize: 13, color: '#0f172a' }}>Import History</div>
            {importHistory.map((batch, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 16px', borderBottom: '1px solid #f8fafc', fontSize: 12 }}>
                <span style={{ color: '#334155' }}>{new Date(batch.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
                <span style={{ color: '#64748b' }}>{batch.fileCount} files · {batch.classified} enriched · {batch.templates} templates</span>
              </div>
            ))}
          </div>
        )}
      </>
    );
  }

  // ============================================================================
  // BATCH PANEL
  // ============================================================================

  private renderBatchPanel(): React.ReactElement {
    const { showBatchPanel, batchTemplateId, batchCategory, batchRisk, selectedIds, fastTrackTemplates } = this.state;
    const templateOptions: IDropdownOption[] = [{ key: '', text: '— No template —' }, ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))];

    return (
      <StyledPanel isOpen={showBatchPanel} onDismiss={() => this.setState({ showBatchPanel: false })}
        headerText={`Batch Assign (${selectedIds.size} selected)`} type={PanelType.smallFixedFar}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            <PrimaryButton text="Apply" disabled={!batchTemplateId && !batchCategory && !batchRisk}
              onClick={() => {
                if (batchTemplateId) { for (const id of selectedIds) this.applyTemplate(id, parseInt(batchTemplateId)); }
                else {
                  for (const id of selectedIds) {
                    const updates: Partial<IBulkImportItem> = {};
                    if (batchCategory) updates.category = batchCategory;
                    if (batchRisk) updates.risk = batchRisk;
                    updates.status = 'enriched';
                    this.updateItem(id, updates);
                  }
                  this.log(`Metadata applied to ${selectedIds.size} files`, 'success');
                }
                this.setState({ showBatchPanel: false, batchTemplateId: '', batchCategory: '', batchRisk: '', selectedIds: new Set() });
              }}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showBatchPanel: false })} styles={{ root: { borderRadius: 4 } }} />
          </Stack>
        )} isFooterAtBottom={true}>
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
          <Text style={{ fontSize: 13, color: '#64748b' }}>Apply a template or set metadata for {selectedIds.size} files.</Text>
          <div style={{ background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 4, padding: 14 }}>
            <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 13, display: 'block', marginBottom: 6 }}>Fast Track Template</Text>
            <Dropdown selectedKey={batchTemplateId} options={templateOptions} onChange={(_, opt) => this.setState({ batchTemplateId: String(opt?.key || '') })} placeholder="Select..." styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          </div>
          <Text style={{ fontSize: 11, color: '#94a3b8', textAlign: 'center' }}>— or set fields —</Text>
          <Dropdown label="Category" selectedKey={batchCategory} options={CATEGORY_OPTIONS} onChange={(_, opt) => this.setState({ batchCategory: String(opt?.key || '') })} styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          <Dropdown label="Risk" selectedKey={batchRisk} options={RISK_OPTIONS} onChange={(_, opt) => this.setState({ batchRisk: String(opt?.key || '') })} styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
        </Stack>
      </StyledPanel>
    );
  }
}
