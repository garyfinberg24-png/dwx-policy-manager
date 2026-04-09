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
import { createDialogManager } from '../../../hooks/useDialog';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { tc } from '../../../utils/themeColors';

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
  enrichSortBy: string; enrichFilterCat: string;
  importHistory: Array<{ date: string; fileCount: number; classified: number; templates: number }>;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const MAX_FILES = 50;
const MAX_FILE_SIZE = 25 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ['.docx', '.pdf', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.rtf', '.txt'];
const SESSION_KEY = 'pm_bulk_upload_state';

/** Map of file extensions to expected MIME types for cross-validation */
const EXPECTED_MIME_MAP: Record<string, string> = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.html': 'text/html',
};

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
  private dialogManager = createDialogManager();

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
      enrichSortBy: 'title', enrichFilterCat: 'All',
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

  /**
   * Extract text from uploaded file — improved extraction for AI classification.
   * DOCX: Decompresses ZIP, parses word/document.xml for all text runs + headings
   * PPTX: Decompresses ZIP, parses slide XML for all text
   * PDF: Extracts text streams between BT/ET markers
   * Plain text: Direct read
   * Target: up to 8000 chars for richer LLM context
   */
  private async extractTextFromFile(file: File): Promise<string> {
    const MAX_CHARS = 8000;
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    try {
      // Plain text formats — direct read
      if (['txt', 'rtf', 'csv', 'md', 'html', 'htm'].includes(ext)) {
        const text = await file.text();
        // Strip HTML tags if present
        const clean = text.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();
        return clean.substring(0, MAX_CHARS);
      }

      // Office XML formats (DOCX, PPTX, XLSX) — ZIP-based extraction
      if (['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'].includes(ext)) {
        const buffer = await file.arrayBuffer();
        const bytes = new Uint8Array(buffer);

        // Find ZIP local file headers and extract XML content
        const xmlParts: string[] = [];
        const textDecoder = new TextDecoder('utf-8', { fatal: false });

        // Look for PK\x03\x04 ZIP signature and extract XML entries
        for (let i = 0; i < bytes.length - 4; i++) {
          if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
            // Parse local file header
            const fnLen = bytes[i + 26] | (bytes[i + 27] << 8);
            const extraLen = bytes[i + 28] | (bytes[i + 29] << 8);
            const fnStart = i + 30;
            const fn = textDecoder.decode(bytes.slice(fnStart, fnStart + fnLen));

            // Only process content XML files
            const isContent = fn === 'word/document.xml' || fn.startsWith('ppt/slides/slide') ||
              fn.startsWith('xl/sharedStrings') || fn === 'xl/worksheets/sheet1.xml';
            if (isContent) {
              const dataStart = fnStart + fnLen + extraLen;
              // Compression method: 0=stored, 8=deflate
              const compMethod = bytes[i + 8] | (bytes[i + 9] << 8);
              const compSize = bytes[i + 18] | (bytes[i + 19] << 8) | (bytes[i + 20] << 16) | (bytes[i + 21] << 24);

              if (compMethod === 0 && compSize > 0) {
                // Stored (uncompressed) — read directly
                const xml = textDecoder.decode(bytes.slice(dataStart, dataStart + compSize));
                xmlParts.push(xml);
              } else if (compMethod === 8 && compSize > 0) {
                // Deflate — try DecompressionStream if available
                try {
                  const compressed = bytes.slice(dataStart, dataStart + compSize);
                  const ds = new (globalThis as any).DecompressionStream('raw');
                  const writer = ds.writable.getWriter();
                  writer.write(compressed);
                  writer.close();
                  const reader = ds.readable.getReader();
                  const chunks: Uint8Array[] = [];
                  let done = false;
                  while (!done) {
                    const result = await reader.read();
                    if (result.value) chunks.push(result.value);
                    done = result.done;
                  }
                  const totalLen = chunks.reduce((s, c) => s + c.length, 0);
                  const merged = new Uint8Array(totalLen);
                  let offset = 0;
                  for (const c of chunks) { merged.set(c, offset); offset += c.length; }
                  xmlParts.push(textDecoder.decode(merged));
                } catch {
                  // DecompressionStream not available — fall back to raw regex
                  const rawSlice = textDecoder.decode(bytes.slice(dataStart, Math.min(dataStart + 500000, bytes.length)));
                  xmlParts.push(rawSlice);
                }
              }
            }
          }
        }

        // Extract text from XML — get all text runs, headings, and paragraph content
        const allXml = xmlParts.join('\n');
        const textParts: string[] = [];

        // Word: <w:t>, <w:t xml:space="preserve">
        // PowerPoint: <a:t>
        // Excel shared strings: <t>
        const textRe = /<(?:w:|a:)?t[^>]*>([^<]+)<\/(?:w:|a:)?t>/g;
        let match: RegExpExecArray | null;
        while ((match = textRe.exec(allXml)) !== null && textParts.length < 2000) {
          const text = match[1].trim();
          if (text.length > 0) textParts.push(text);
        }

        // Also extract any text between paragraph markers that the <t> regex missed
        const paraRe = />([^<]{10,})</g;
        while ((match = paraRe.exec(allXml)) !== null && textParts.length < 2500) {
          const text = match[1].trim().replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
          if (text.length > 10 && !textParts.includes(text)) textParts.push(text);
        }

        if (textParts.length > 3) {
          return textParts.join(' ').replace(/\s{2,}/g, ' ').substring(0, MAX_CHARS);
        }
      }

      // PDF — extract text between BT/ET markers and parenthesized strings
      if (ext === 'pdf') {
        const buffer = await file.arrayBuffer();
        const bytes = new Uint8Array(buffer);
        const raw = new TextDecoder('latin1').decode(bytes.slice(0, Math.min(bytes.length, 1048576)));

        // Method 1: Extract parenthesized strings from text objects (Tj, TJ operators)
        const textParts: string[] = [];
        const tjRe = /\(([^)]{2,})\)\s*Tj/g;
        let match: RegExpExecArray | null;
        while ((match = tjRe.exec(raw)) !== null && textParts.length < 2000) {
          const t = match[1].replace(/\\([nrt\\()])/g, (_, c) => ({ n: '\n', r: '\r', t: '\t', '\\': '\\', '(': '(', ')': ')' }[c] || c));
          if (t.trim().length > 1) textParts.push(t.trim());
        }

        // Method 2: TJ array operator
        const tjArrayRe = /\[([^\]]*)\]\s*TJ/g;
        while ((match = tjArrayRe.exec(raw)) !== null && textParts.length < 2500) {
          const inner = match[1];
          const strRe = /\(([^)]+)\)/g;
          let sm: RegExpExecArray | null;
          const words: string[] = [];
          while ((sm = strRe.exec(inner)) !== null) { if (sm[1].trim()) words.push(sm[1].trim()); }
          if (words.length > 0) textParts.push(words.join(''));
        }

        // Method 3: Fallback — printable ASCII sequences
        if (textParts.length < 20) {
          const ascii = raw.replace(/[^\x20-\x7E\n]/g, ' ').replace(/\s{3,}/g, ' ');
          const sentences = ascii.split(/\s{2,}/).filter(s => s.length > 25 && /[a-zA-Z]{3,}/.test(s));
          for (const s of sentences.slice(0, 500)) {
            if (!textParts.includes(s)) textParts.push(s);
          }
        }

        if (textParts.length > 3) {
          return textParts.join(' ').replace(/\s{2,}/g, ' ').substring(0, MAX_CHARS);
        }
      }
    } catch (err) {
      console.warn('[BulkUpload] Text extraction failed:', err);
    }
    // Fallback: filename only
    return `Document filename: ${file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')}`;
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
      // MIME type cross-validation: reject files where extension and MIME don't match
      const expectedMime = EXPECTED_MIME_MAP[ext];
      if (expectedMime && file.type && file.type !== expectedMime) {
        this.log(`${file.name}: MIME type mismatch (expected ${expectedMime}, got ${file.type})`, 'warning');
        continue;
      }
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
        const title = item.title || item.fileName.replace(/\.[^.]+$/, '');

        // 1. Create policy draft first so we have an spId for per-policy folder
        const result = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.add({ Title: title, PolicyStatus: 'Draft' });
        const spId = result?.data?.Id || result?.data?.id;

        // 2. Try per-policy subfolder in PM_PolicySourceDocuments/{spId}, fall back to BulkImports/
        let targetFolderUrl = folderUrl;
        if (spId) {
          const policyFolderUrl = `${siteRelUrl}/${PM_LISTS.POLICY_SOURCE_DOCUMENTS}/${spId}`;
          try {
            const xf = new XMLHttpRequest();
            xf.open('POST', `${siteUrl}/_api/web/folders`, false);
            xf.setRequestHeader('Accept', 'application/json; odata=verbose');
            xf.setRequestHeader('Content-Type', 'application/json; odata=verbose');
            xf.setRequestHeader('X-RequestDigest', digest);
            xf.send(JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': policyFolderUrl }));
            if (xf.status >= 200 && xf.status < 400) {
              targetFolderUrl = policyFolderUrl;
            }
          } catch {
            // Folder creation failed — use BulkImports/ fallback
          }
        }

        // 3. Upload file to the target folder
        const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${targetFolderUrl}')/Files/AddUsingPath(decodedurl='${encodeURIComponent(safeName)}',overwrite=true)`;
        const docUrl: string = await new Promise((res, rej) => {
          const x = new XMLHttpRequest(); x.open('POST', endpoint, true);
          x.setRequestHeader('Accept', 'application/json; odata=verbose'); x.setRequestHeader('Content-Type', 'application/octet-stream'); x.setRequestHeader('X-RequestDigest', digest);
          x.responseType = 'json'; x.onload = () => x.status >= 200 && x.status < 300 ? res(x.response?.d?.ServerRelativeUrl || '') : rej(new Error(`${x.status}`)); x.onerror = () => rej(new Error('Network')); x.send(new Uint8Array(buf));
        });

        // 4. Update policy item with metadata and document URL
        if (spId) { try { const updateData: Record<string, unknown> = { PolicyName: title, CreationMethod: 'BulkImport' }; if (docUrl) updateData.DocumentURL = docUrl; await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(spId).update(updateData); } catch { /* CreationMethod column may not exist */ } }
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

    // Write bulk import audit trail to PM_PolicyAuditLog
    try {
      const currentUserEmail = this.props.context?.pageContext?.user?.email || '';
      const summary = toUpload.map(i => ({
        fileName: i.fileName,
        category: i.category || 'Unclassified',
        status: this.state.imports.find(x => x.id === i.id)?.status || 'unknown',
      }));
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_AUDIT_LOG).items.add({
        Title: `Bulk Import \u2014 ${succeeded} policies`,
        AuditAction: 'BulkImport',
        ActionDescription: JSON.stringify({ fileCount: toUpload.length, succeeded, failed: toUpload.length - succeeded, files: summary }).substring(0, 4000),
        PerformedByEmail: currentUserEmail,
        ActionDate: new Date().toISOString(),
        EntityType: 'Policy',
      });
    } catch {
      // PM_PolicyAuditLog may not exist — non-blocking
    }
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
          const contentSnippet = content.length > 100 ? content.substring(0, 8000) : '';
          const classificationPrompt = `You are an expert policy analyst for a corporate Policy Management system. Analyze this document and extract structured metadata.

DOCUMENT FILENAME: "${item.fileName}"
${contentSnippet ? `\nDOCUMENT CONTENT (first ${contentSnippet.length} characters):\n"""\n${contentSnippet}\n"""` : '\n(No document content extracted — classify from filename only)'}

TASK: Extract the following metadata by analyzing the document content. Use the content to make informed decisions — do not guess blindly from the filename alone.

Think step by step:
1. Read the document content carefully
2. Identify the subject matter and regulatory context
3. Determine the appropriate category and risk level based on content
4. Extract key points and a concise summary

REQUIRED OUTPUT (respond with ONLY this JSON object, no other text):
{
  "title": "Clean, professional policy title (e.g., 'Information Security Policy' not 'InfoSec_Policy_v2_FINAL')",
  "category": "EXACTLY ONE OF: IT Security | HR | Compliance | Data Protection | Health & Safety | Finance | Legal | Operations | Governance | Other",
  "risk": "EXACTLY ONE OF: Critical | High | Medium | Low | Informational",
  "departments": "Comma-separated list of target departments (e.g., 'All Employees' or 'IT, Engineering' or 'Finance, Legal')",
  "summary": "2-3 sentence summary of what this policy covers and why it matters",
  "readTimeframe": "EXACTLY ONE OF: Immediate | Day 1 | Day 3 | Week 1 | Week 2 | Month 1",
  "keyPoints": ["Key point 1", "Key point 2", "Key point 3"],
  "regulatoryReferences": "Any regulatory frameworks mentioned (e.g., 'POPIA, GDPR' or 'ISO 27001' or 'None detected')",
  "reviewFrequency": "EXACTLY ONE OF: Annual | Biannual | Quarterly | As Needed",
  "requiresAcknowledgement": true or false (true if the policy requires staff to formally acknowledge they have read it)
}

CLASSIFICATION GUIDANCE:
- Critical risk: Legal/regulatory obligations, data breaches, health/safety hazards
- High risk: Security policies, financial controls, compliance requirements
- Medium risk: Operational procedures, HR policies, general guidelines
- Low risk: Best practices, recommendations, informational guides
- Immediate/Day 1: Critical safety or compliance. Week 1: Standard policies. Month 1: Reference material.`;

          try {
            const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), 45000);
            const resp = await fetch(functionUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ mode: 'author-assist', message: classificationPrompt, conversationHistory: [], policyContext: [], userRole: 'Author' }),
              signal: ctrl.signal });
            clearTimeout(t);
            if (resp.ok) {
              const d = await resp.json();
              const raw = d?.message || d?.response || d?.content || '';
              try { const m = raw.match(/\{[\s\S]*\}/); if (m) suggestions = JSON.parse(m[0]); } catch { /* JSON parse failed */ }
            }
          } catch { /* AI call failed — will fall through to heuristic */ }
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
          keyPoints: Array.isArray(suggestions.keyPoints) ? suggestions.keyPoints : [],
          regulatoryReferences: suggestions.regulatoryReferences || '',
          reviewFrequency: suggestions.reviewFrequency || 'Annual',
          requiresAcknowledgement: suggestions.requiresAcknowledgement ?? true,
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
      <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr', minHeight: 'calc(100vh - 180px)', background: '#fff', borderRadius: 10, overflow: 'hidden', border: '1px solid #e2e8f0', margin: '0 auto', maxWidth: 1600 }}>
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
                  style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 16px', cursor: clickable ? 'pointer' : 'default', borderLeft: active ? `3px solid ${tc.primary}` : '3px solid transparent', background: active ? tc.primaryLighter : 'transparent' }}>
                  <div style={{ width: 26, height: 26, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, flexShrink: 0, background: done ? tc.primary : active ? tc.primaryLighter : '#fff', color: done ? '#fff' : active ? tc.primary : '#94a3b8', border: `2px solid ${done ? tc.primary : active ? tc.primary : '#e2e8f0'}` }}>
                    {done ? <Icon iconName="CheckMark" style={{ fontSize: 10 }} /> : step.num}
                  </div>
                  <div><div style={{ fontWeight: active ? 600 : 500, color: active ? tc.primary : done ? '#0f172a' : '#475569', fontSize: 12 }}>{step.title}</div><div style={{ fontSize: 9, color: '#94a3b8' }}>{step.desc}</div></div>
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
                styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}>
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
          style={{ border: `2px dashed ${dragOver ? tc.primary : '#cbd5e1'}`, borderRadius: 10, padding: '36px 24px', textAlign: 'center', cursor: 'pointer', background: dragOver ? tc.primaryLighter : '#fff', marginBottom: 16 }}>
          <input ref={this._fileInputRef} type="file" multiple accept={ALLOWED_EXTENSIONS.join(',')} onChange={this.handleFileSelect} style={{ display: 'none' }} />
          <svg viewBox="0 0 24 24" fill="none" width="32" height="32" style={{ color: dragOver ? tc.primary : '#94a3b8', marginBottom: 8 }}><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /><path d="M17 8l-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /><path d="M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /></svg>
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
            styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }} />
        )}
      </>
    );
  }

  // ============================================================================
  // STEP 2: REVIEW
  // ============================================================================

  private renderStep2_Review(): React.ReactElement {
    const { imports, searchQuery, selectedIds, filterType, groupBy } = this.state;
    let uploaded = imports.filter(i => !['pending', 'failed'].includes(i.status));

    // Filter
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); uploaded = uploaded.filter(i => i.title.toLowerCase().includes(q) || i.fileName.toLowerCase().includes(q)); }
    if (filterType !== 'All') uploaded = uploaded.filter(i => i.fileType === `.${filterType.toLowerCase()}`);

    // Sort
    const sortKey = this.state.enrichSortBy || 'title';
    uploaded = [...uploaded].sort((a, b) => {
      if (sortKey === 'title') return a.title.localeCompare(b.title);
      if (sortKey === 'type') return a.fileType.localeCompare(b.fileType);
      if (sortKey === 'metadata') return (b.hasExistingMetadata ? 1 : 0) - (a.hasExistingMetadata ? 1 : 0);
      return 0;
    });

    // Group
    const groups: Array<{ key: string; items: typeof uploaded }> = [];
    if (groupBy === 'None' || !groupBy) {
      groups.push({ key: '', items: uploaded });
    } else {
      const map = new Map<string, typeof uploaded>();
      for (const item of uploaded) {
        const key = groupBy === 'Type' ? item.fileType.replace('.', '').toUpperCase() : groupBy === 'Metadata' ? (item.hasExistingMetadata ? 'Has Metadata' : 'No Metadata') : groupBy === 'Author' ? (item.existingMetadata?.author || 'Unknown') : '';
        if (!map.has(key)) map.set(key, []);
        map.get(key)!.push(item);
      }
      for (const [key, items] of Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]))) groups.push({ key, items });
    }

    const filtered = uploaded;
    const allSelected = filtered.length > 0 && filtered.every(i => selectedIds.has(i.id));

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Review Uploaded Documents</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 14px' }}>Check existing metadata, edit titles. Select files for the next step.</p>

        <div style={{ display: 'flex', gap: 8, marginBottom: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <SearchBox placeholder="Search..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 180, borderRadius: 4 }, field: { borderRadius: 4 } }} />
          <Dropdown selectedKey={filterType} options={[{ key: 'All', text: 'All Types' }, { key: 'DOCX', text: 'DOCX' }, { key: 'PDF', text: 'PDF' }, { key: 'XLSX', text: 'XLSX' }, { key: 'PPTX', text: 'PPTX' }]}
            onChange={(_, opt) => this.setState({ filterType: String(opt?.key || 'All') })}
            styles={{ root: { width: 110 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 12 }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
          <Dropdown selectedKey={sortKey} options={[{ key: 'title', text: 'Sort: Title' }, { key: 'type', text: 'Sort: Type' }, { key: 'metadata', text: 'Sort: Metadata' }]}
            onChange={(_, opt) => this.setState({ enrichSortBy: String(opt?.key || 'title') })}
            styles={{ root: { width: 120 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 12 }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
          <Dropdown selectedKey={groupBy || 'None'} options={[{ key: 'None', text: 'No grouping' }, { key: 'Type', text: 'Group: Type' }, { key: 'Metadata', text: 'Group: Metadata' }, { key: 'Author', text: 'Group: Author' }]}
            onChange={(_, opt) => this.setState({ groupBy: String(opt?.key || 'None') })}
            styles={{ root: { width: 130 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 12 }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 11, color: '#94a3b8' }}>{selectedIds.size} selected · {filtered.length} files</span>
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '32px 1fr 90px 90px 90px 80px', padding: '8px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" checked={allSelected} onChange={() => { if (allSelected) this.setState({ selectedIds: new Set() }); else this.setState({ selectedIds: new Set(filtered.map(i => i.id)) }); }} /></div>
            <div>Title / File</div><div>Author</div><div>Category</div><div>Keywords</div><div>Metadata</div>
          </div>
          {groups.map(group => (
            <React.Fragment key={group.key || '__all'}>
              {group.key && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 14px', background: tc.primaryLighter, borderBottom: '1px solid #e2e8f0', cursor: 'pointer' }}
                  onClick={() => { const ids = group.items.map(i => i.id); const allGrp = ids.every(id => selectedIds.has(id)); const next = new Set(selectedIds); if (allGrp) ids.forEach(id => next.delete(id)); else ids.forEach(id => next.add(id)); this.setState({ selectedIds: next }); }}>
                  <input type="checkbox" checked={group.items.every(i => selectedIds.has(i.id))} readOnly />
                  <span style={{ fontSize: 12, fontWeight: 700, color: tc.primary }}>{group.key}</span>
                  <span style={{ fontSize: 11, color: '#94a3b8' }}>({group.items.length})</span>
                </div>
              )}
              {group.items.map(item => {
            const meta = item.existingMetadata || {};
            return (
              <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '32px 1fr 90px 90px 90px 80px', padding: '8px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: selectedIds.has(item.id) ? tc.primaryLighter : '#fff' }}>
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
            </React.Fragment>
          ))}
        </div>
      </>
    );
  }

  // ============================================================================
  // STEP 3: ENRICH METADATA (merged Classify + Templates)
  // ============================================================================

  private renderStep3_Enrich(siteUrl: string): React.ReactElement {
    const { imports, classifying, classifyProgress, selectedIds, fastTrackTemplates, enrichSortBy, enrichFilterCat, searchQuery } = this.state;
    let enrichable = imports.filter(i => !['pending', 'failed'].includes(i.status));
    const templateOptions: IDropdownOption[] = [{ key: '', text: '— No template —' }, ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))];
    const selectedCount = selectedIds.size;
    const riskColor = (r: string) => r === 'Critical' ? '#dc2626' : r === 'High' ? '#d97706' : r === 'Medium' ? tc.primary : r === 'Low' ? '#059669' : '#94a3b8';

    // Filter
    if (enrichFilterCat !== 'All') enrichable = enrichable.filter(i => i.category === enrichFilterCat);
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); enrichable = enrichable.filter(i => i.title.toLowerCase().includes(q) || i.fileName.toLowerCase().includes(q)); }
    // Sort
    enrichable = [...enrichable].sort((a, b) => {
      switch (enrichSortBy) {
        case 'title': return a.title.localeCompare(b.title);
        case 'category': return (a.category || '').localeCompare(b.category || '');
        case 'risk': { const order = ['Critical', 'High', 'Medium', 'Low', 'Informational', '']; return order.indexOf(a.risk) - order.indexOf(b.risk); }
        case 'status': return (a.status || '').localeCompare(b.status || '');
        default: return 0;
      }
    });

    // Fill-down helper
    const fillDown = (field: 'category' | 'risk' | 'department' | 'readTimeframe', value: string): void => {
      const updated = this.state.imports.map(i => selectedIds.has(i.id) ? { ...i, [field]: value, status: i.status === 'uploaded' ? 'enriched' as ImportStatus : i.status } : i);
      this.setState({ imports: updated });
      this.log(`Fill down: ${field} = "${value}" → ${selectedIds.size} files`, 'success');
    };

    return (
      <>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10 }}>
          <div>
            <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Enrich Metadata</h2>
            <p style={{ fontSize: 13, color: '#64748b', margin: 0 }}>Edit directly, use AI, apply templates, or fill down for batches. All optional.</p>
          </div>
        </div>

        {/* Toolbar */}
        <div style={{ display: 'flex', gap: 6, marginBottom: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <SearchBox placeholder="Search..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 180, borderRadius: 4 }, field: { borderRadius: 4 } }} />
          <Dropdown selectedKey={enrichFilterCat} options={[{ key: 'All', text: 'All Categories' }, ...CATEGORY_OPTIONS.filter(o => o.key)]}
            onChange={(_, opt) => this.setState({ enrichFilterCat: String(opt?.key || 'All') })}
            styles={{ root: { width: 130 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 12 }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
          <Dropdown selectedKey={enrichSortBy} options={[{ key: 'title', text: 'Sort: Title' }, { key: 'category', text: 'Sort: Category' }, { key: 'risk', text: 'Sort: Risk' }, { key: 'status', text: 'Sort: Status' }]}
            onChange={(_, opt) => this.setState({ enrichSortBy: String(opt?.key || 'title') })}
            styles={{ root: { width: 120 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 12 }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
          <div style={{ flex: 1 }} />
          {selectedCount > 0 && !classifying && (
            <>
              <PrimaryButton text={`AI Classify (${selectedCount})`} iconProps={{ iconName: 'Processing' }}
                onClick={() => this.classifySelected()}
                styles={{ root: { background: '#7c3aed', borderColor: '#7c3aed', borderRadius: 4, fontSize: 12, height: 30 }, rootHovered: { background: '#6d28d9', borderColor: '#6d28d9' } }} />
              <DefaultButton text={`Batch Assign (${selectedCount})`} iconProps={{ iconName: 'Tag' }}
                onClick={() => this.setState({ showBatchPanel: true })}
                styles={{ root: { borderRadius: 4, fontSize: 12, height: 30 } }} />
              {/* Fill-down quick actions */}
              <Dropdown placeholder="Fill Category ↓" options={CATEGORY_OPTIONS.filter(o => o.key)}
                onChange={(_, opt) => { if (opt?.key) fillDown('category', String(opt.key)); }}
                styles={{ root: { width: 130 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 11, color: tc.primary, borderColor: '#99f6e4' }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
              <Dropdown placeholder="Fill Risk ↓" options={RISK_OPTIONS.filter(o => o.key)}
                onChange={(_, opt) => { if (opt?.key) fillDown('risk', String(opt.key)); }}
                styles={{ root: { width: 110 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 11, color: tc.primary, borderColor: '#99f6e4' }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
              <Dropdown placeholder="Fill Template ↓" selectedKey="" options={[{ key: '', text: 'Fill Template ↓' }, ...templateOptions.filter(o => o.key)]}
                onChange={(_, opt) => {
                  if (opt?.key) {
                    const templateId = parseInt(String(opt.key));
                    const ids = Array.from(selectedIds);
                    ids.forEach(id => this.applyTemplate(id, templateId));
                    this.log(`Template applied to ${ids.length} files`, 'success');
                  }
                }}
                styles={{ root: { width: 150 }, title: { borderRadius: 4, height: 30, lineHeight: '28px', fontSize: 11, color: '#059669', borderColor: '#bbf7d0' }, caretDownWrapper: { height: 30, lineHeight: '30px' } }} />
            </>
          )}
          <span style={{ fontSize: 11, color: '#94a3b8' }}>{selectedCount > 0 ? `${selectedCount} sel` : ''} · {enrichable.length} files</span>
        </div>

        {classifying && <ProgressIndicator label={`Classifying... ${classifyProgress}%`} percentComplete={classifyProgress / 100} style={{ marginBottom: 10 }} />}

        {/* Editable enrichment table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '32px 1.5fr 140px 100px 130px 140px 60px 40px', padding: '8px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', gap: 6 }}>
            <div><input type="checkbox" onChange={(e) => { if ((e.target as HTMLInputElement).checked) this.setState({ selectedIds: new Set(enrichable.map(i => i.id)) }); else this.setState({ selectedIds: new Set() }); }} /></div>
            <div>Policy Title</div><div>Category</div><div>Risk</div><div>Department</div><div>Template</div><div>Status</div><div></div>
          </div>
          {enrichable.map(item => {
            const isClassifying = item.status === 'classifying';
            return (
              <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '32px 1.5fr 140px 100px 130px 140px 60px 40px', padding: '8px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: isClassifying ? '#faf5ff' : selectedIds.has(item.id) ? tc.primaryLighter : '#fff', opacity: isClassifying ? 0.7 : 1, gap: 6 }}>
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

                {/* Department — editable dropdown */}
                <div>
                  <Dropdown selectedKey={item.department || 'All Employees'} disabled={isClassifying}
                    options={[{ key: 'All Employees', text: 'All Employees' }, { key: 'IT', text: 'IT' }, { key: 'HR', text: 'HR' }, { key: 'Finance', text: 'Finance' }, { key: 'Legal', text: 'Legal' }, { key: 'Operations', text: 'Operations' }, { key: 'Compliance', text: 'Compliance' }, { key: 'Executive', text: 'Executive' }, { key: 'Engineering', text: 'Engineering' }, { key: 'Sales', text: 'Sales' }]}
                    onChange={(_, opt) => this.updateItem(item.id, { department: String(opt?.key || 'All Employees'), status: item.status === 'uploaded' ? 'enriched' : item.status })}
                    styles={{ root: { minWidth: 0 }, title: { fontSize: 11, height: 26, lineHeight: '24px', borderRadius: 4, borderColor: '#e2e8f0' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }} />
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
                  {item.spId && <IconButton iconProps={{ iconName: 'Edit' }} title="Open in Policy Builder"
                    onClick={async () => { const ok = await this.dialogManager.showConfirm(`Open "${item.title}" in Policy Builder?\n\nYou will leave the Bulk Upload wizard. Your progress is saved.`, { title: 'Open in Builder', confirmText: 'Open', cancelText: 'Cancel' }); if (ok) window.location.href = `${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${item.spId}`; }}
                    styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 12, color: tc.primary } }} />}
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
    const { imports, activityLog, importHistory, selectedIds } = this.state;
    const processedItems = imports.filter(i => !['pending', 'failed'].includes(i.status));
    const uploaded = processedItems.length;
    const enriched = imports.filter(i => ['classified', 'enriched'].includes(i.status)).length;
    const templated = imports.filter(i => !!i.templateId).length;
    const failed = imports.filter(i => i.status === 'failed').length;
    const selectedCount = Array.from(selectedIds).filter(id => processedItems.some(p => p.id === id)).length;
    const allSelected = processedItems.length > 0 && processedItems.every(i => selectedIds.has(i.id));

    // Batch action helper
    const batchUpdateStatus = async (status: string, label: string, itemFilter: (i: IBulkImportItem) => boolean, extraFields?: Record<string, unknown>) => {
      const targets = itemFilter === null ? processedItems.filter(i => i.spId) : processedItems.filter(i => i.spId && itemFilter(i));
      const desc = targets.length === uploaded ? 'all' : `${targets.length} selected`;
      const confirmed = await this.dialogManager.showConfirm(`Are you sure you want to ${label.toLowerCase()} ${desc} polic${targets.length !== 1 ? 'ies' : 'y'}?${status === 'Published' ? '\n\nThis will make them visible to their target audience.' : ''}`, { title: label, confirmText: label, cancelText: 'Cancel' });
      if (!confirmed) return;
      let count = 0;
      for (const item of targets) {
        try { await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({ PolicyStatus: status, ...extraFields }); count++; } catch { /* */ }
      }
      this.log(`${label}: ${count} policies`, 'success');
      this.setState({ successMessage: `${count} policies ${label.toLowerCase()}.`, selectedIds: new Set() });
    };

    return (
      <>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Import Summary</h2>
        <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 16px' }}>Review your imported policies. Select files to approve or publish individually, or use batch actions.</p>

        {/* KPI cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 16 }}>
          {[{ l: 'Uploaded', v: uploaded, c: '#2563eb' }, { l: 'Enriched', v: enriched, c: '#7c3aed' }, { l: 'With Template', v: templated, c: '#059669' }, { l: 'Failed', v: failed, c: failed > 0 ? '#dc2626' : '#94a3b8' }].map(k =>
            <div key={k.l} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.c}`, padding: '12px 14px', textAlign: 'center' }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: k.c }}>{k.v}</div>
              <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.l}</div>
            </div>
          )}
        </div>

        {/* Action buttons */}
        <div style={{ display: 'flex', gap: 6, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
          <PrimaryButton text="Open Drafts & Pipeline" iconProps={{ iconName: 'ViewAll' }} href={`${siteUrl}/SitePages/PolicyAuthor.aspx`}
            styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 4, fontSize: 12, height: 32 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }} />
          <div style={{ width: 1, height: 24, background: '#e2e8f0', margin: '0 4px' }} />
          {selectedCount > 0 ? (
            <>
              <DefaultButton text={`Submit Selected (${selectedCount})`} iconProps={{ iconName: 'Send' }}
                onClick={() => batchUpdateStatus('In Review', 'Submitted for review', (i) => selectedIds.has(i.id))}
                styles={{ root: { borderRadius: 4, fontSize: 12, height: 32, color: '#2563eb', borderColor: '#93c5fd' }, rootHovered: { background: '#eff6ff' } }} />
              <DefaultButton text={`Publish Selected (${selectedCount})`} iconProps={{ iconName: 'PublishContent' }}
                onClick={() => batchUpdateStatus('Published', 'Published', (i) => selectedIds.has(i.id), { IsActive: true })}
                styles={{ root: { borderRadius: 4, fontSize: 12, height: 32, color: '#059669', borderColor: '#bbf7d0' }, rootHovered: { background: '#f0fdf4' } }} />
            </>
          ) : null}
          <DefaultButton text="Submit All for Review" iconProps={{ iconName: 'Send' }} disabled={uploaded === 0}
            onClick={() => batchUpdateStatus('In Review', 'Submitted for review', null)}
            styles={{ root: { borderRadius: 4, fontSize: 12, height: 32, color: '#2563eb', borderColor: '#dbeafe' } }} />
          <DefaultButton text="Publish All" iconProps={{ iconName: 'PublishContent' }} disabled={uploaded === 0}
            onClick={() => batchUpdateStatus('Published', 'Published', null, { IsActive: true })}
            styles={{ root: { borderRadius: 4, fontSize: 12, height: 32, color: '#059669', borderColor: '#d1fae5' } }} />
          <div style={{ flex: 1 }} />
          <DefaultButton text="New Import" iconProps={{ iconName: 'Add' }}
            onClick={() => {
              const batch = { date: new Date().toISOString(), fileCount: imports.length, classified: enriched, templates: templated };
              this.setState({ wizardStep: 1, completedSteps: new Set(), imports: [], selectedIds: new Set(), activityLog: [], importHistory: [batch, ...this.state.importHistory.slice(0, 19)] });
              sessionStorage.removeItem(SESSION_KEY);
            }}
            styles={{ root: { borderRadius: 4, fontSize: 12, height: 32 } }} />
        </div>

        {/* Processed Files Table with checkboxes */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0' }}>
            <span style={{ fontWeight: 600, fontSize: 13, color: '#0f172a' }}>Processed Files ({processedItems.length})</span>
            {selectedCount > 0 && <span style={{ fontSize: 12, color: tc.primary, fontWeight: 600 }}>{selectedCount} selected</span>}
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '32px 1fr 110px 80px 100px 140px 70px', padding: '6px 16px', background: '#fafafa', borderBottom: '1px solid #f1f5f9', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div><input type="checkbox" checked={allSelected} onChange={() => { if (allSelected) this.setState({ selectedIds: new Set() }); else this.setState({ selectedIds: new Set(processedItems.map(i => i.id)) }); }} /></div>
            <div>Policy Title</div><div>Category</div><div>Risk</div><div>Department</div><div>Template</div><div>Status</div>
          </div>
          <div style={{ maxHeight: 400, overflowY: 'auto' }}>
            {processedItems.map(item => {
              const rc = item.risk === 'Critical' ? '#dc2626' : item.risk === 'High' ? '#d97706' : item.risk === 'Medium' ? tc.primary : '#059669';
              const sc = item.status === 'failed' ? '#dc2626' : ['classified', 'enriched'].includes(item.status) ? '#059669' : '#2563eb';
              const sl = item.status === 'failed' ? 'Failed' : ['classified', 'enriched'].includes(item.status) ? 'Ready' : 'Uploaded';
              const isSelected = selectedIds.has(item.id);
              return (
                <div key={item.id} style={{ display: 'grid', gridTemplateColumns: '32px 1fr 110px 80px 100px 140px 70px', padding: '8px 16px', borderBottom: '1px solid #f8fafc', alignItems: 'center', fontSize: 12, background: isSelected ? tc.primaryLighter : '#fff' }}>
                  <div><input type="checkbox" checked={isSelected} onChange={() => { const n = new Set(selectedIds); if (n.has(item.id)) n.delete(item.id); else n.add(item.id); this.setState({ selectedIds: n }); }} /></div>
                  <div><div style={{ fontWeight: 600, color: '#0f172a', fontSize: 13 }}>{item.title}</div><div style={{ fontSize: 10, color: '#cbd5e1' }}>{item.fileName}</div></div>
                  <div>{item.category ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f5f3ff', color: '#7c3aed' }}>{item.category}</span> : <span style={{ color: '#cbd5e1' }}>—</span>}</div>
                  <div>{item.risk ? <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: `${rc}10`, color: rc }}>{item.risk}</span> : <span style={{ color: '#cbd5e1' }}>—</span>}</div>
                  <div style={{ fontSize: 11, color: '#475569' }}>{item.department || '—'}</div>
                  <div style={{ fontSize: 11, color: item.templateName ? '#059669' : '#cbd5e1' }}>{item.templateName || '—'}</div>
                  <div><span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: `${sc}10`, color: sc }}>{sl}</span></div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Collapsible Activity Log */}
        <details style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 }}>
          <summary style={{ padding: '10px 16px', background: '#f8fafc', cursor: 'pointer', fontWeight: 600, fontSize: 12, color: '#64748b', userSelect: 'none' }}>
            Activity Log ({activityLog.length} entries)
          </summary>
          <div style={{ maxHeight: 200, overflowY: 'auto' }}>
            {activityLog.map((e, i) => (
              <div key={i} style={{ display: 'flex', gap: 8, padding: '5px 16px', borderBottom: '1px solid #f8fafc', fontSize: 11 }}>
                <span style={{ color: '#94a3b8', flexShrink: 0, minWidth: 55 }}>{e.time.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                <span style={{ width: 6, height: 6, borderRadius: '50%', marginTop: 4, flexShrink: 0, background: e.type === 'success' ? '#059669' : e.type === 'error' ? '#dc2626' : e.type === 'warning' ? '#d97706' : '#94a3b8' }} />
                <span style={{ color: e.type === 'error' ? '#dc2626' : '#334155' }}>{e.message}</span>
              </div>
            ))}
          </div>
        </details>

        {/* Import History — collapsible */}
        {importHistory.length > 0 && (
          <details style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <summary style={{ padding: '10px 16px', background: '#f8fafc', cursor: 'pointer', fontWeight: 600, fontSize: 12, color: '#64748b', userSelect: 'none' }}>
              Import History ({importHistory.length} batches)
            </summary>
            {importHistory.map((batch, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 16px', borderBottom: '1px solid #f8fafc', fontSize: 12 }}>
                <span style={{ color: '#334155' }}>{new Date(batch.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
                <span style={{ color: '#64748b' }}>{batch.fileCount} files · {batch.classified} enriched · {batch.templates} templates</span>
              </div>
            ))}
          </details>
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
              styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showBatchPanel: false })} styles={{ root: { borderRadius: 4 } }} />
          </Stack>
        )} isFooterAtBottom={true}>
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
          <Text style={{ fontSize: 13, color: '#64748b' }}>Apply a template or set metadata for {selectedIds.size} files.</Text>
          <div style={{ background: tc.primaryLighter, border: `1px solid ${tc.primaryLight}`, borderRadius: 4, padding: 14 }}>
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
