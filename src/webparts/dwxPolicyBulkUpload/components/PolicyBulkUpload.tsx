// @ts-nocheck
import * as React from 'react';
import { IPolicyBulkUploadProps } from './IPolicyBulkUploadProps';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, IconButton, SearchBox, Dropdown, IDropdownOption,
  ProgressIndicator, Checkbox, Toggle, TextField
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

type UploadPhase = 'upload' | 'classify' | 'review';

type ImportStatus = 'uploaded' | 'classifying' | 'classified' | 'metadata-complete' | 'ready' | 'failed';

interface IFastTrackTemplate {
  Id: number;
  Title: string;
  ProfileName: string;
  PolicyCategory: string;
  ComplianceRisk: string;
  ReadTimeframe: string;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
  TargetDepartments: string;
}

type MatchConfidence = 'Strong' | 'Likely' | 'Possible' | 'None';

interface IBulkImportItem {
  id: string;  // temp ID until SP item created
  spId?: number; // SP list item ID once created
  fileName: string;
  fileSize: number;
  fileType: string;
  file?: File;
  documentUrl?: string;
  status: ImportStatus;
  // AI-extracted metadata
  extractedTitle?: string;
  suggestedCategory?: string;
  suggestedRisk?: string;
  suggestedDepartments?: string[];
  suggestedSummary?: string;
  suggestedKeyPoints?: string[];
  suggestedReadTimeframe?: string;
  suggestedRequiresAck?: boolean;
  // Template matching
  matchedTemplateId?: number;
  matchedTemplateName?: string;
  matchConfidence?: MatchConfidence;
  // User-confirmed metadata (editable in classify table)
  confirmedTitle?: string;
  confirmedCategory?: string;
  confirmedRisk?: string;
  confirmedDepartments?: string[];
  confirmedTemplateId?: number;
  templateApplied?: boolean;
  // Selection
  selected?: boolean;
  error?: string;
  classifyProgress?: number;
}

interface IPolicyBulkUploadState {
  loading: boolean;
  detectedRole: PolicyManagerRole | null;
  phase: UploadPhase;
  imports: IBulkImportItem[];
  uploading: boolean;
  uploadProgress: number;
  classifying: boolean;
  classifyProgress: number;
  selectedIds: Set<string>;
  searchQuery: string;
  statusFilter: 'All' | ImportStatus;
  batchCategory: string;
  batchRisk: string;
  batchTemplateId: string;
  showBatchPanel: boolean;
  successMessage: string;
  errorMessage: string;
  dragOver: boolean;
  fastTrackTemplates: IFastTrackTemplate[];
  templatesLoaded: boolean;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const MAX_FILES = 50;
const MAX_FILE_SIZE = 25 * 1024 * 1024; // 25MB
const ALLOWED_EXTENSIONS = ['.docx', '.pdf', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.rtf', '.txt'];

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' },
  { key: 'IT Security', text: 'IT Security' },
  { key: 'HR', text: 'Human Resources' },
  { key: 'Compliance', text: 'Compliance' },
  { key: 'Data Protection', text: 'Data Protection' },
  { key: 'Health & Safety', text: 'Health & Safety' },
  { key: 'Finance', text: 'Finance' },
  { key: 'Legal', text: 'Legal' },
  { key: 'Operations', text: 'Operations' },
  { key: 'Governance', text: 'Governance' },
  { key: 'Other', text: 'Other' }
];

const RISK_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(select)' },
  { key: 'Critical', text: 'Critical' },
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' },
  { key: 'Informational', text: 'Informational' }
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
    this.state = {
      loading: true,
      detectedRole: null,
      phase: 'upload',
      imports: [],
      uploading: false,
      uploadProgress: 0,
      classifying: false,
      classifyProgress: 0,
      selectedIds: new Set(),
      searchQuery: '',
      statusFilter: 'All',
      batchCategory: '',
      batchRisk: '',
      batchTemplateId: '',
      showBatchPanel: false,
      successMessage: '',
      errorMessage: '',
      dragOver: false,
      fastTrackTemplates: [],
      templatesLoaded: false
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this.detectRole();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  private async detectRole(): Promise<void> {
    try {
      const roleService = new RoleDetectionService(this.props.sp);
      const userRoles = await roleService.getCurrentUserRoles();
      const role = userRoles && userRoles.length > 0 ? getHighestPolicyRole(userRoles) : PolicyManagerRole.User;
      if (this._isMounted) this.setState({ detectedRole: role, loading: false });
    } catch {
      // Fallback: assume Author role so the page loads
      if (this._isMounted) this.setState({ detectedRole: PolicyManagerRole.Author, loading: false });
    }
  }

  // ============================================================================
  // FAST TRACK TEMPLATE LOADING + MATCHING
  // ============================================================================

  private async loadFastTrackTemplates(): Promise<void> {
    if (this.state.templatesLoaded) return;
    try {
      const items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
        .items.filter('IsActive ne 0')
        .select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk', 'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz', 'TargetDepartments')
        .orderBy('Title').top(100)();
      const templates: IFastTrackTemplate[] = items.map((t: any) => ({
        Id: t.Id, Title: t.Title || t.ProfileName, ProfileName: t.ProfileName || t.Title,
        PolicyCategory: t.PolicyCategory || '', ComplianceRisk: t.ComplianceRisk || 'Medium',
        ReadTimeframe: t.ReadTimeframe || 'Week 1',
        RequiresAcknowledgement: t.RequiresAcknowledgement !== false,
        RequiresQuiz: t.RequiresQuiz || false,
        TargetDepartments: t.TargetDepartments || ''
      }));
      if (this._isMounted) this.setState({ fastTrackTemplates: templates, templatesLoaded: true });
    } catch {
      if (this._isMounted) this.setState({ templatesLoaded: true });
    }
  }

  private matchTemplate(item: IBulkImportItem): { templateId: number; templateName: string; confidence: MatchConfidence } | null {
    const { fastTrackTemplates } = this.state;
    if (fastTrackTemplates.length === 0) return null;

    const category = (item.suggestedCategory || '').toLowerCase();
    const risk = (item.suggestedRisk || '').toLowerCase();
    const depts = (item.suggestedDepartments || []).map(d => d.toLowerCase());

    let bestMatch: IFastTrackTemplate | null = null;
    let bestScore = 0;

    for (const tmpl of fastTrackTemplates) {
      let score = 0;
      // Category match (strongest signal)
      if (tmpl.PolicyCategory.toLowerCase() === category) score += 3;
      // Risk match
      if (tmpl.ComplianceRisk.toLowerCase() === risk) score += 2;
      // Department overlap
      const tmplDepts = (tmpl.TargetDepartments || '').split(';').map(d => d.trim().toLowerCase()).filter(Boolean);
      if (tmplDepts.length > 0 && depts.some(d => tmplDepts.includes(d))) score += 1;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = tmpl;
      }
    }

    if (!bestMatch || bestScore === 0) return null;

    const confidence: MatchConfidence = bestScore >= 5 ? 'Strong' : bestScore >= 3 ? 'Likely' : 'Possible';
    return { templateId: bestMatch.Id, templateName: bestMatch.Title || bestMatch.ProfileName, confidence };
  }

  private async applyTemplateToItem(itemId: string, templateId: number): Promise<void> {
    const template = this.state.fastTrackTemplates.find(t => t.Id === templateId);
    if (!template) return;

    const item = this.state.imports.find(i => i.id === itemId);
    if (!item?.spId) return;

    // Update SP list item with template metadata
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({
        PolicyCategory: template.PolicyCategory,
        ComplianceRisk: template.ComplianceRisk,
        ReadTimeframe: template.ReadTimeframe,
        RequiresAcknowledgement: template.RequiresAcknowledgement,
        RequiresQuiz: template.RequiresQuiz,
        Departments: template.TargetDepartments
      });
    } catch { /* best-effort — column may not exist */ }

    // Update local state
    const updated = this.state.imports.map(i =>
      i.id === itemId ? {
        ...i,
        confirmedCategory: template.PolicyCategory,
        confirmedRisk: template.ComplianceRisk,
        confirmedTemplateId: templateId,
        templateApplied: true,
        status: 'metadata-complete' as ImportStatus
      } : i
    );
    this.setState({ imports: updated });
  }

  private async applyTemplateToSelected(templateId: number): Promise<void> {
    const { selectedIds } = this.state;
    for (const id of selectedIds) {
      await this.applyTemplateToItem(id, templateId);
    }
    this.setState({
      selectedIds: new Set(),
      successMessage: `Template applied to ${selectedIds.size} polic${selectedIds.size !== 1 ? 'ies' : 'y'}.`
    });
    setTimeout(() => this.setState({ successMessage: '' }), 4000);
  }

  private async acceptAllSuggestions(): Promise<void> {
    const { imports } = this.state;
    const classifiedItems = imports.filter(i => i.status === 'classified' && i.matchedTemplateId);
    for (const item of classifiedItems) {
      if (item.matchedTemplateId) {
        await this.applyTemplateToItem(item.id, item.matchedTemplateId);
      }
    }
    this.setState({
      successMessage: `Accepted AI suggestions for ${classifiedItems.length} polic${classifiedItems.length !== 1 ? 'ies' : 'y'}.`
    });
    setTimeout(() => this.setState({ successMessage: '' }), 4000);
  }

  // ============================================================================
  // PHASE 1: FILE UPLOAD
  // ============================================================================

  private handleFileDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    this.setState({ dragOver: false });
    const files = Array.from(e.dataTransfer.files);
    this.processFiles(files);
  };

  private handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const files = Array.from(e.target.files || []);
    this.processFiles(files);
    if (this._fileInputRef.current) this._fileInputRef.current.value = '';
  };

  private processFiles(files: File[]): void {
    const { imports } = this.state;
    const remaining = MAX_FILES - imports.length;
    if (remaining <= 0) {
      this.setState({ errorMessage: `Maximum ${MAX_FILES} files per batch.` });
      return;
    }

    const validFiles: IBulkImportItem[] = [];
    const errors: string[] = [];

    for (const file of files.slice(0, remaining)) {
      const ext = '.' + file.name.split('.').pop()?.toLowerCase();
      if (!ALLOWED_EXTENSIONS.includes(ext)) {
        errors.push(`${file.name}: unsupported format (${ext})`);
        continue;
      }
      if (file.size > MAX_FILE_SIZE) {
        errors.push(`${file.name}: exceeds 25MB limit`);
        continue;
      }
      // Check duplicate
      if (imports.some(i => i.fileName === file.name)) {
        errors.push(`${file.name}: already added`);
        continue;
      }
      validFiles.push({
        id: `import_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        fileName: file.name,
        fileSize: file.size,
        fileType: ext,
        file,
        status: 'uploaded',
        policyTitle: file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')
      });
    }

    this.setState({
      imports: [...imports, ...validFiles],
      errorMessage: errors.length > 0 ? errors.join('; ') : ''
    });
  }

  private removeImport = (id: string): void => {
    this.setState(prev => ({
      imports: prev.imports.filter(i => i.id !== id),
      selectedIds: (() => { const s = new Set(prev.selectedIds); s.delete(id); return s; })()
    }));
  };

  private async uploadToSharePoint(): Promise<void> {
    const { imports } = this.state;
    const toUpload = imports.filter(i => i.status === 'uploaded' && i.file);
    if (toUpload.length === 0) return;

    this.setState({ uploading: true, uploadProgress: 0 });
    const currentUser = await this.props.sp.web.currentUser();
    let processed = 0;

    // Get site URL for REST API calls
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
    const siteServerRelativeUrl = this.props.context?.pageContext?.web?.serverRelativeUrl || '/sites/PolicyManager';

    // Ensure BulkImports folder exists (once for entire batch)
    try {
      const folderRelativeUrl = `${siteServerRelativeUrl}/${PM_LISTS.POLICY_SOURCE_DOCUMENTS}/BulkImports`;
      try {
        await this.props.sp.web.getFolderByServerRelativePath(folderRelativeUrl)();
      } catch {
        // Create folder via REST API to avoid PnP serialization issues
        try {
          const ctx = await this.props.sp.web.select('Url')();
          await fetch(`${siteUrl}/_api/web/folders`, {
            method: 'POST',
            headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': (document.getElementById('__REQUESTDIGEST') as HTMLInputElement)?.value || '' },
            body: JSON.stringify({ ServerRelativeUrl: folderRelativeUrl })
          });
        } catch { /* folder may already exist */ }
      }
    } catch { /* non-blocking */ }

    for (const item of toUpload) {
      try {
        // Upload file using SPFx HttpClient (proper auth context)
        const fileBuffer = await item.file.arrayBuffer();
        const folderRelativeUrl = `${siteServerRelativeUrl}/${PM_LISTS.POLICY_SOURCE_DOCUMENTS}/BulkImports`;
        const safeFileName = item.fileName.replace(/[#%&*:<>?\/\\{|}~]/g, '_');

        // Use SPFx context's spHttpClient for proper auth
        const { SPHttpClient } = await import('@microsoft/sp-http');
        const uploadEndpoint = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folderRelativeUrl}')/Files/AddUsingPath(decodedurl='${encodeURIComponent(safeFileName)}',overwrite=true)`;

        const uploadResponse = await this.props.context.spHttpClient.post(
          uploadEndpoint,
          SPHttpClient.configurations.v1,
          {
            body: fileBuffer,
            headers: { 'Accept': 'application/json;odata=nometadata' }
          }
        );

        let docUrl = '';
        if (uploadResponse.ok) {
          const uploadData = await uploadResponse.json();
          docUrl = uploadData.ServerRelativeUrl || '';
        } else {
          const errText = await uploadResponse.text();
          throw new Error(`Upload failed: ${uploadResponse.status} — ${errText.substring(0, 200)}`);
        }

        // Create Draft policy stub in PM_Policies
        const policyTitle = item.policyTitle || item.fileName.replace(/\.[^.]+$/, '');
        const policyData: Record<string, unknown> = {
          Title: policyTitle,
          PolicyName: policyTitle,
          PolicyStatus: 'Draft',
          DocumentURL: docUrl,
          PolicyCategory: '',
          ComplianceRisk: 'Medium'
        };
        // Optional columns — only include if they exist (avoids 400 on missing columns)
        try {
          policyData.DocumentFormat = item.fileType.replace('.', '').toUpperCase();
        } catch { /* */ }
        const policyResult = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.add(policyData);
        const spId = policyResult?.data?.Id || policyResult?.data?.id;

        // Update local state
        const updated = this.state.imports.map(i =>
          i.id === item.id ? { ...i, spId, documentUrl: docUrl, file: undefined } : i
        );
        processed++;
        if (this._isMounted) {
          this.setState({ imports: updated, uploadProgress: Math.round((processed / toUpload.length) * 100) });
        }
      } catch (err) {
        console.error(`[BulkUpload] Failed to upload ${item.fileName}:`, err);
        const updated = this.state.imports.map(i =>
          i.id === item.id ? { ...i, status: 'failed' as ImportStatus, error: 'Upload failed' } : i
        );
        processed++;
        if (this._isMounted) this.setState({ imports: updated, uploadProgress: Math.round((processed / toUpload.length) * 100) });
      }
    }

    if (this._isMounted) {
      this.setState({ uploading: false, phase: 'classify', successMessage: `${processed} file${processed !== 1 ? 's' : ''} uploaded successfully.` });
      setTimeout(() => this.setState({ successMessage: '' }), 4000);
    }
  }

  // ============================================================================
  // PHASE 2: AI CLASSIFICATION
  // ============================================================================

  private async classifyWithAI(): Promise<void> {
    // Re-read state to get fresh data after upload
    const imports = this.state.imports;
    const toClassify = imports.filter(i => i.status === 'uploaded' && i.spId);
    console.log(`[BulkUpload] classifyWithAI: ${imports.length} total imports, ${toClassify.length} ready to classify`);
    if (toClassify.length === 0) {
      console.warn('[BulkUpload] No items to classify — all items:', imports.map(i => ({ id: i.id, status: i.status, spId: i.spId })));
      this.setState({ errorMessage: 'No uploaded items found to classify. Try uploading first.' });
      return;
    }

    this.setState({ classifying: true, classifyProgress: 0 });

    // Load Fast Track Templates for matching
    await this.loadFastTrackTemplates();

    // Get AI chat function URL from config or localStorage
    let functionUrl = '';
    try {
      const configItems = await this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Integration.AI.Chat.FunctionUrl'")
        .select('ConfigValue').top(1)();
      functionUrl = configItems[0]?.ConfigValue || '';
    } catch { /* fallback */ }
    if (!functionUrl) {
      functionUrl = localStorage.getItem('PM_AI_ChatFunctionUrl') || '';
    }

    console.log(`[BulkUpload] AI function URL: ${functionUrl ? 'configured' : 'NOT configured — using heuristic fallback'}`);
    let processed = 0;

    for (const item of toClassify) {
      try {
        console.log(`[BulkUpload] Classifying: ${item.fileName} (spId: ${item.spId})`);
        // Mark as classifying
        let updated = this.state.imports.map(i =>
          i.id === item.id ? { ...i, status: 'classifying' as ImportStatus } : i
        );
        if (this._isMounted) this.setState({ imports: updated });

        let suggestions: any = {};

        if (functionUrl) {
          // Call AI to classify the document
          const controller = new AbortController();
          const timeout = setTimeout(() => controller.abort(), 30000);

          try {
            const response = await fetch(functionUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                mode: 'author-assist',
                message: `Classify this policy document. Based on the filename "${item.fileName}" and any title hints, extract and suggest:
1. Title — a clean, professional policy title (e.g., "Information Security Policy" not the filename)
2. PolicyCategory (one of: IT Security, HR, Compliance, Data Protection, Health & Safety, Finance, Legal, Operations, Governance, Other)
3. ComplianceRisk (one of: Critical, High, Medium, Low, Informational)
4. Departments (comma-separated, e.g., "IT, All Employees")
5. Summary (1-2 sentences describing the policy)
6. KeyPoints (3-5 bullet points of what the policy covers)
7. ReadTimeframe (one of: Immediate, Day 1, Day 3, Week 1, Week 2, Month 1)
8. RequiresAcknowledgement (true/false — true if it's a compliance/mandatory policy)

Respond ONLY with a JSON object with keys: title, category, risk, departments, summary, keyPoints, readTimeframe, requiresAck`,
                history: [],
                context: []
              }),
              signal: controller.signal
            });
            clearTimeout(timeout);

            if (response.ok) {
              const data = await response.json();
              const content = data?.response || data?.content || '';
              // Try to parse JSON from the response
              try {
                const jsonMatch = content.match(/\{[\s\S]*\}/);
                if (jsonMatch) {
                  suggestions = JSON.parse(jsonMatch[0]);
                }
              } catch { /* AI response wasn't valid JSON — use heuristic fallback */ }
            }
          } catch { /* AI call failed — use heuristic fallback */ }
        }

        // Heuristic fallback if AI didn't return valid suggestions
        if (!suggestions.category) {
          suggestions = this.heuristicClassify(item.fileName, item.policyTitle || '');
        }

        // Build the classified item
        const classifiedItem: Partial<IBulkImportItem> = {
          status: 'classified' as ImportStatus,
          extractedTitle: suggestions.title || item.policyTitle || item.fileName.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' '),
          suggestedCategory: suggestions.category || 'Other',
          suggestedRisk: suggestions.risk || 'Medium',
          suggestedDepartments: Array.isArray(suggestions.departments) ? suggestions.departments : (suggestions.departments || '').split(',').map((d: string) => d.trim()).filter(Boolean),
          suggestedSummary: suggestions.summary || '',
          suggestedKeyPoints: Array.isArray(suggestions.keyPoints) ? suggestions.keyPoints : [],
          suggestedReadTimeframe: suggestions.readTimeframe || 'Week 1',
          suggestedRequiresAck: suggestions.requiresAck === true || suggestions.requiresAck === 'true',
          // Pre-fill confirmed with suggestions
          confirmedTitle: suggestions.title || item.policyTitle || item.fileName.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' '),
          confirmedCategory: suggestions.category || 'Other',
          confirmedRisk: suggestions.risk || 'Medium',
          confirmedDepartments: Array.isArray(suggestions.departments) ? suggestions.departments : (suggestions.departments || '').split(',').map((d: string) => d.trim()).filter(Boolean),
        };

        // Match to a Fast Track Template
        const tempItem = { ...item, ...classifiedItem } as IBulkImportItem;
        const templateMatch = this.matchTemplate(tempItem);
        if (templateMatch) {
          classifiedItem.matchedTemplateId = templateMatch.templateId;
          classifiedItem.matchedTemplateName = templateMatch.templateName;
          classifiedItem.matchConfidence = templateMatch.confidence;
          classifiedItem.confirmedTemplateId = templateMatch.templateId;
        }

        updated = this.state.imports.map(i =>
          i.id === item.id ? {
            ...i,
            ...classifiedItem
          } : i
        );
        processed++;
        if (this._isMounted) {
          this.setState({ imports: updated, classifyProgress: Math.round((processed / toClassify.length) * 100) });
        }
      } catch (err) {
        console.error(`[BulkUpload] Classification failed for ${item.fileName}:`, err);
        const updated = this.state.imports.map(i =>
          i.id === item.id ? { ...i, status: 'classified' as ImportStatus, suggestedCategory: 'Other', suggestedRisk: 'Medium', confirmedCategory: 'Other', confirmedRisk: 'Medium' } : i
        );
        processed++;
        if (this._isMounted) this.setState({ imports: updated, classifyProgress: Math.round((processed / toClassify.length) * 100) });
      }
    }

    if (this._isMounted) {
      this.setState({ classifying: false, phase: 'review', successMessage: `${processed} polic${processed !== 1 ? 'ies' : 'y'} classified.` });
      setTimeout(() => this.setState({ successMessage: '' }), 4000);
    }
  }

  private heuristicClassify(fileName: string, title: string): any {
    const text = (fileName + ' ' + title).toLowerCase();
    let category = 'Other';
    let risk = 'Medium';
    if (/security|cyber|access|password|firewall/i.test(text)) { category = 'IT Security'; risk = 'High'; }
    else if (/hr|human|employee|leave|conduct|harassment|disciplin/i.test(text)) { category = 'HR'; risk = 'Medium'; }
    else if (/compliance|regulat|audit|sox|iso/i.test(text)) { category = 'Compliance'; risk = 'High'; }
    else if (/data|privacy|gdpr|pii|personal/i.test(text)) { category = 'Data Protection'; risk = 'Critical'; }
    else if (/health|safety|incident|hazard|ohs|whs/i.test(text)) { category = 'Health & Safety'; risk = 'High'; }
    else if (/financ|expense|procurement|budget|travel/i.test(text)) { category = 'Finance'; risk = 'Medium'; }
    else if (/legal|contract|intellectual|nda|confidential/i.test(text)) { category = 'Legal'; risk = 'High'; }
    else if (/operat|process|procedure|workflow|standard/i.test(text)) { category = 'Operations'; risk = 'Low'; }
    else if (/govern|board|ethic|whistleblow|conflict/i.test(text)) { category = 'Governance'; risk = 'High'; }
    // Clean title from filename
    const cleanTitle = (fileName + '').replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    return { title: cleanTitle, category, risk, departments: ['All Employees'], summary: '', keyPoints: [], readTimeframe: 'Week 1', requiresAck: ['Critical', 'High'].includes(risk) };
  }

  // ============================================================================
  // PHASE 3: REVIEW & BATCH ASSIGN
  // ============================================================================

  private applyBatchMetadata = async (): Promise<void> => {
    const { selectedIds, imports, batchCategory, batchRisk } = this.state;
    if (selectedIds.size === 0) return;

    const updated = imports.map(i => {
      if (!selectedIds.has(i.id)) return i;
      return {
        ...i,
        confirmedCategory: batchCategory || i.confirmedCategory,
        confirmedRisk: batchRisk || i.confirmedRisk,
        status: 'metadata-complete' as ImportStatus
      };
    });

    // Update SP items
    for (const item of updated) {
      if (selectedIds.has(item.id) && item.spId) {
        try {
          await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({
            PolicyCategory: item.confirmedCategory || '',
            ComplianceRisk: item.confirmedRisk || 'Medium',
            PolicyDescription: item.suggestedSummary || ''
          });
        } catch { /* per-item — continue on failure */ }
      }
    }

    this.setState({ imports: updated, showBatchPanel: false, selectedIds: new Set(), successMessage: `Metadata applied to ${selectedIds.size} polic${selectedIds.size !== 1 ? 'ies' : 'y'}.` });
    setTimeout(() => this.setState({ successMessage: '' }), 4000);
  };

  private acceptAISuggestions = async (itemId: string): Promise<void> => {
    const updated = this.state.imports.map(i => {
      if (i.id !== itemId) return i;
      return { ...i, status: 'metadata-complete' as ImportStatus };
    });
    const item = updated.find(i => i.id === itemId);
    if (item?.spId) {
      try {
        await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(item.spId).update({
          PolicyCategory: item.confirmedCategory || item.suggestedCategory || '',
          ComplianceRisk: item.confirmedRisk || item.suggestedRisk || 'Medium',
          PolicyDescription: item.suggestedSummary || ''
        });
      } catch { /* best-effort */ }
    }
    this.setState({ imports: updated });
  };

  // ============================================================================
  // RENDER
  // ============================================================================

  public render(): React.ReactElement {
    const { detectedRole, loading } = this.state;

    if (loading) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Bulk Upload.">
          <JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp}
            activeNavKey="bulk-upload"
            breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}>
            <div style={{ padding: 60, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading..." /></div>
          </JmlAppLayout>
        </ErrorBoundary>
      );
    }

    if (detectedRole !== null && !hasMinimumRole(detectedRole, PolicyManagerRole.Author)) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Bulk Upload.">
          <JmlAppLayout title="Bulk Upload" context={this.props.context} sp={this.props.sp}
            activeNavKey="bulk-upload"
            breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}>
            <section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}>
              <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} />
              <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>Access Denied</Text>
              <Text variant="medium" block styles={{ root: { color: '#64748b' } }}>Bulk Upload requires an Author role or higher.</Text>
            </section>
          </JmlAppLayout>
        </ErrorBoundary>
      );
    }

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Bulk Upload.">
        <JmlAppLayout title={this.props.title || 'Policy Bulk Upload'} context={this.props.context} sp={this.props.sp}
          activeNavKey="bulk-upload"
          breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Bulk Upload' }]}>
          {this.renderContent()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  private renderContent(): React.ReactElement {
    const { phase, imports, successMessage, errorMessage, uploading, classifying } = this.state;
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    return (
      <section style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        {/* Page Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <h1 style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', margin: '0 0 4px 0' }}>Bulk Upload</h1>
            <p style={{ fontSize: 13, color: '#64748b', margin: 0 }}>Import existing policies with AI-powered classification and metadata assignment</p>
          </div>
        </div>

        {/* Messages */}
        {successMessage && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ successMessage: '' })} style={{ marginBottom: 12 }}>{successMessage}</MessageBar>}
        {errorMessage && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ errorMessage: '' })} style={{ marginBottom: 12 }}>{errorMessage}</MessageBar>}

        {/* Phase Progress Indicator */}
        <div style={{ display: 'flex', gap: 0, marginBottom: 28 }}>
          {[
            { key: 'upload', label: 'Upload', icon: 'CloudUpload', number: 1 },
            { key: 'classify', label: 'AI Classify', icon: 'Processing', number: 2 },
            { key: 'review', label: 'Review & Assign', icon: 'CheckboxComposite', number: 3 }
          ].map((step, i) => {
            const phases: UploadPhase[] = ['upload', 'classify', 'review'];
            const currentIdx = phases.indexOf(phase);
            const stepIdx = phases.indexOf(step.key as UploadPhase);
            const isDone = stepIdx < currentIdx;
            const isCurrent = stepIdx === currentIdx;
            return (
              <React.Fragment key={step.key}>
                <div
                  onClick={() => { if (isDone) this.setState({ phase: step.key as UploadPhase }); }}
                  style={{
                    flex: 1, display: 'flex', alignItems: 'center', gap: 10, padding: '14px 20px',
                    background: isCurrent ? '#f0fdfa' : isDone ? '#fff' : '#fafafa',
                    border: `1px solid ${isCurrent ? '#0d9488' : '#e2e8f0'}`,
                    borderRadius: i === 0 ? '10px 0 0 10px' : i === 2 ? '0 10px 10px 0' : 0,
                    cursor: isDone ? 'pointer' : 'default', transition: 'all 0.2s'
                  }}
                >
                  <div style={{
                    width: 28, height: 28, borderRadius: '50%',
                    background: isDone ? '#0d9488' : isCurrent ? '#0d9488' : '#e2e8f0',
                    color: isDone || isCurrent ? '#fff' : '#94a3b8',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 13, fontWeight: 700, flexShrink: 0
                  }}>
                    {isDone ? '✓' : step.number}
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: isCurrent ? '#0d9488' : isDone ? '#0f172a' : '#94a3b8' }}>{step.label}</div>
                  </div>
                </div>
              </React.Fragment>
            );
          })}
        </div>

        {/* KPI Cards */}
        {imports.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 12, marginBottom: 20 }}>
            {[
              { label: 'Total', value: imports.length, color: '#475569' },
              { label: 'Uploaded', value: imports.filter(i => i.status === 'uploaded').length, color: '#2563eb' },
              { label: 'Classified', value: imports.filter(i => ['classified', 'metadata-complete', 'ready'].includes(i.status)).length, color: '#7c3aed' },
              { label: 'Metadata Done', value: imports.filter(i => ['metadata-complete', 'ready'].includes(i.status)).length, color: '#059669' },
              { label: 'Failed', value: imports.filter(i => i.status === 'failed').length, color: '#dc2626' }
            ].map(kpi => (
              <div key={kpi.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, borderTop: `3px solid ${kpi.color}`, padding: '12px 14px', textAlign: 'center' }}>
                <div style={{ fontSize: 22, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
                <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{kpi.label}</div>
              </div>
            ))}
          </div>
        )}

        {/* Phase Content */}
        {phase === 'upload' && this.renderUploadPhase()}
        {phase === 'classify' && this.renderClassifyPhase()}
        {phase === 'review' && this.renderReviewPhase(siteUrl)}

        {/* Batch Metadata Panel */}
        {this.renderBatchPanel()}
      </section>
    );
  }

  // ---- UPLOAD PHASE ----
  private renderUploadPhase(): React.ReactElement {
    const { imports, uploading, uploadProgress, dragOver } = this.state;

    return (
      <>
        {/* Drag & Drop Zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); this.setState({ dragOver: true }); }}
          onDragLeave={() => this.setState({ dragOver: false })}
          onDrop={this.handleFileDrop}
          onClick={() => this._fileInputRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? '#0d9488' : '#cbd5e1'}`,
            borderRadius: 12, padding: '48px 32px', textAlign: 'center', cursor: 'pointer',
            background: dragOver ? '#f0fdfa' : '#fafafa', transition: 'all 0.2s', marginBottom: 20
          }}
        >
          <input ref={this._fileInputRef} type="file" multiple accept={ALLOWED_EXTENSIONS.join(',')} onChange={this.handleFileSelect} style={{ display: 'none' }} />
          <svg viewBox="0 0 24 24" fill="none" width="40" height="40" style={{ color: dragOver ? '#0d9488' : '#94a3b8', marginBottom: 12 }}>
            <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            <path d="M17 8l-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            <path d="M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
          </svg>
          <div style={{ fontSize: 15, fontWeight: 600, color: '#0f172a', marginBottom: 4 }}>
            {dragOver ? 'Drop files here' : 'Drag & drop policy documents here'}
          </div>
          <div style={{ fontSize: 12, color: '#94a3b8' }}>
            or click to browse — DOCX, PDF, XLSX, PPTX, up to 25MB each — max {MAX_FILES} files
          </div>
        </div>

        {/* File List */}
        {imports.length > 0 && (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <Text style={{ fontWeight: 600, color: '#0f172a' }}>{imports.length} file{imports.length !== 1 ? 's' : ''} selected</Text>
              <DefaultButton text="Clear All" onClick={() => this.setState({ imports: [], selectedIds: new Set() })} styles={{ root: { fontSize: 12, height: 28 } }} />
            </div>
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', marginBottom: 20 }}>
              {imports.map(item => {
                const sizeStr = item.fileSize < 1024 * 1024
                  ? `${Math.round(item.fileSize / 1024)} KB`
                  : `${(item.fileSize / (1024 * 1024)).toFixed(1)} MB`;
                return (
                  <div key={item.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 16px', borderBottom: '1px solid #f1f5f9' }}>
                    <Icon iconName={this.getFileIcon(item.fileType)} styles={{ root: { fontSize: 20, color: '#0d9488' } }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{item.fileName}</div>
                      <div style={{ fontSize: 11, color: '#94a3b8' }}>{sizeStr} • {item.fileType.toUpperCase()}</div>
                    </div>
                    {item.status === 'failed' && <span style={{ fontSize: 11, color: '#dc2626', fontWeight: 600 }}>{item.error || 'Failed'}</span>}
                    <IconButton iconProps={{ iconName: 'Cancel' }} title="Remove" onClick={() => this.removeImport(item.id)} styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 12, color: '#dc2626' } }} />
                  </div>
                );
              })}
            </div>

            {uploading && <ProgressIndicator label={`Uploading... ${uploadProgress}%`} percentComplete={uploadProgress / 100} style={{ marginBottom: 16 }} />}

            <div style={{ display: 'flex', gap: 8 }}>
              <PrimaryButton
                text={uploading ? 'Uploading...' : `Upload ${imports.length} File${imports.length !== 1 ? 's' : ''}`}
                iconProps={{ iconName: 'CloudUpload' }}
                disabled={uploading || imports.length === 0}
                onClick={() => this.uploadToSharePoint()}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              />
              <DefaultButton
                text="Upload & Classify"
                iconProps={{ iconName: 'Processing' }}
                disabled={uploading || imports.length === 0}
                onClick={async () => {
                  await this.uploadToSharePoint();
                  // Small delay to let React state flush after upload
                  await new Promise(r => setTimeout(r, 500));
                  await this.classifyWithAI();
                }}
                styles={{ root: { borderRadius: 4 } }}
              />
            </div>
          </>
        )}
      </>
    );
  }

  // ---- CLASSIFY PHASE ----
  private renderClassifyPhase(): React.ReactElement {
    const { imports, classifying, classifyProgress, selectedIds, fastTrackTemplates, searchQuery } = this.state;
    const unclassified = imports.filter(i => i.status === 'uploaded' && i.spId);
    const classified = imports.filter(i => ['classified', 'metadata-complete'].includes(i.status));
    const withTemplateMatch = classified.filter(i => i.matchedTemplateId);

    // Template dropdown options
    const templateOptions: IDropdownOption[] = [
      { key: '', text: '— No template —' },
      ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title || t.ProfileName }))
    ];

    // Update a field on an import item
    const updateItem = (id: string, field: string, value: any): void => {
      const updated = this.state.imports.map(i => i.id === id ? { ...i, [field]: value } : i);
      this.setState({ imports: updated });
    };

    // Confidence badge colours
    const confidenceStyle = (c?: MatchConfidence) => {
      switch (c) {
        case 'Strong': return { bg: '#f0fdf4', color: '#059669' };
        case 'Likely': return { bg: '#eff6ff', color: '#2563eb' };
        case 'Possible': return { bg: '#fef3c7', color: '#d97706' };
        default: return { bg: '#f1f5f9', color: '#94a3b8' };
      }
    };

    // Filtered items for display
    let displayItems = imports.filter(i => i.suggestedCategory || i.status === 'classifying');
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      displayItems = displayItems.filter(i => (i.confirmedTitle || i.extractedTitle || i.fileName).toLowerCase().includes(q));
    }

    const allDisplaySelected = displayItems.length > 0 && displayItems.every(i => selectedIds.has(i.id));

    return (
      <>
        {/* Header + Actions */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <div>
            <h2 style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', margin: 0 }}>AI Classification & Template Matching</h2>
            <p style={{ fontSize: 12, color: '#64748b', margin: '4px 0 0' }}>
              {unclassified.length > 0 ? `${unclassified.length} awaiting classification` : ''}{unclassified.length > 0 && classified.length > 0 ? ' · ' : ''}{classified.length > 0 ? `${classified.length} classified` : ''}{withTemplateMatch.length > 0 ? ` · ${withTemplateMatch.length} template matches` : ''}
            </p>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {unclassified.length > 0 && (
              <PrimaryButton
                text={classifying ? 'Classifying...' : `Classify ${unclassified.length}`}
                iconProps={{ iconName: 'Processing' }}
                disabled={classifying}
                onClick={() => this.classifyWithAI()}
                styles={{ root: { background: '#7c3aed', borderColor: '#7c3aed', borderRadius: 4, fontSize: 12, height: 30 }, rootHovered: { background: '#6d28d9', borderColor: '#6d28d9' } }}
              />
            )}
            {withTemplateMatch.length > 0 && (
              <DefaultButton
                text={`Accept All Matches (${withTemplateMatch.length})`}
                iconProps={{ iconName: 'Accept' }}
                onClick={() => this.acceptAllSuggestions()}
                styles={{ root: { fontSize: 12, height: 30, borderRadius: 4, color: '#059669', borderColor: '#bbf7d0' }, rootHovered: { background: '#f0fdf4', borderColor: '#059669' } }}
              />
            )}
            {selectedIds.size > 0 && (
              <DefaultButton
                text={`Apply Template (${selectedIds.size})`}
                iconProps={{ iconName: 'Tag' }}
                onClick={() => this.setState({ showBatchPanel: true })}
                styles={{ root: { fontSize: 12, height: 30, borderRadius: 4 } }}
              />
            )}
            {classified.length > 0 && (
              <PrimaryButton text="Continue to Review" iconProps={{ iconName: 'Forward' }} onClick={() => this.setState({ phase: 'review' })}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4, fontSize: 12, height: 30 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              />
            )}
          </div>
        </div>

        {/* Search */}
        <div style={{ marginBottom: 12 }}>
          <SearchBox placeholder="Search by title or filename..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 280 } }} />
        </div>

        {classifying && <ProgressIndicator label={`Classifying... ${classifyProgress}%`} percentComplete={classifyProgress / 100} style={{ marginBottom: 16 }} />}

        {/* Classification Table */}
        {displayItems.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            {/* Table Header */}
            <div style={{
              display: 'grid', gridTemplateColumns: '36px 1fr 130px 90px 110px 180px 80px 70px',
              padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
              fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b'
            }}>
              <div><input type="checkbox" checked={allDisplaySelected} onChange={() => {
                if (allDisplaySelected) this.setState({ selectedIds: new Set() });
                else this.setState({ selectedIds: new Set(displayItems.map(i => i.id)) });
              }} /></div>
              <div>Extracted Title</div>
              <div>Category</div>
              <div>Risk</div>
              <div>Department</div>
              <div>Fast Track Template</div>
              <div>Confidence</div>
              <div>Actions</div>
            </div>

            {/* Table Rows */}
            {displayItems.map(item => {
              const isSelected = selectedIds.has(item.id);
              const cStyle = confidenceStyle(item.matchConfidence);
              const isClassifying = item.status === 'classifying';

              return (
                <div key={item.id} style={{
                  display: 'grid', gridTemplateColumns: '36px 1fr 130px 90px 110px 180px 80px 70px',
                  padding: '10px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center',
                  background: isSelected ? '#f0fdfa' : item.templateApplied ? '#f0fdf4' : '#fff',
                  opacity: isClassifying ? 0.6 : 1
                }}>
                  <div><input type="checkbox" checked={isSelected} onChange={() => {
                    const next = new Set(selectedIds);
                    if (next.has(item.id)) next.delete(item.id); else next.add(item.id);
                    this.setState({ selectedIds: next });
                  }} disabled={isClassifying} /></div>

                  {/* Extracted Title — editable */}
                  <div>
                    {isClassifying ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Spinner size={SpinnerSize.xSmall} />
                        <span style={{ fontSize: 12, color: '#94a3b8' }}>Classifying...</span>
                      </div>
                    ) : (
                      <>
                        <input
                          type="text"
                          value={item.confirmedTitle || item.extractedTitle || ''}
                          onChange={(e) => updateItem(item.id, 'confirmedTitle', (e.target as HTMLInputElement).value)}
                          style={{ width: '100%', border: 'none', background: 'transparent', fontSize: 13, fontWeight: 600, color: '#0f172a', outline: 'none', padding: '2px 0' }}
                          title="Click to edit title"
                        />
                        <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {item.fileName}
                          {item.suggestedSummary && <span> · {item.suggestedSummary.substring(0, 60)}...</span>}
                        </div>
                      </>
                    )}
                  </div>

                  {/* Category — dropdown */}
                  <div>
                    <Dropdown
                      selectedKey={item.confirmedCategory || item.suggestedCategory || ''}
                      options={CATEGORY_OPTIONS}
                      onChange={(_, opt) => updateItem(item.id, 'confirmedCategory', opt?.key || '')}
                      styles={{ root: { minWidth: 0 }, title: { fontSize: 12, height: 26, lineHeight: '24px', borderColor: '#e2e8f0' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }}
                      disabled={isClassifying}
                    />
                  </div>

                  {/* Risk — dropdown */}
                  <div>
                    <Dropdown
                      selectedKey={item.confirmedRisk || item.suggestedRisk || ''}
                      options={RISK_OPTIONS}
                      onChange={(_, opt) => updateItem(item.id, 'confirmedRisk', opt?.key || '')}
                      styles={{ root: { minWidth: 0 }, title: { fontSize: 12, height: 26, lineHeight: '24px', borderColor: '#e2e8f0' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }}
                      disabled={isClassifying}
                    />
                  </div>

                  {/* Department */}
                  <div style={{ fontSize: 12, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {(item.confirmedDepartments || item.suggestedDepartments || []).join(', ') || '-'}
                  </div>

                  {/* Fast Track Template — lookup dropdown */}
                  <div>
                    <Dropdown
                      selectedKey={item.confirmedTemplateId ? String(item.confirmedTemplateId) : (item.matchedTemplateId ? String(item.matchedTemplateId) : '')}
                      options={templateOptions}
                      onChange={(_, opt) => updateItem(item.id, 'confirmedTemplateId', opt?.key ? parseInt(String(opt.key)) : undefined)}
                      styles={{ root: { minWidth: 0 }, title: { fontSize: 12, height: 26, lineHeight: '24px', borderColor: item.matchedTemplateId ? '#bbf7d0' : '#e2e8f0', background: item.templateApplied ? '#f0fdf4' : '#fff' }, caretDownWrapper: { height: 26, lineHeight: '26px' } }}
                      disabled={isClassifying}
                    />
                  </div>

                  {/* Confidence badge */}
                  <div>
                    {item.matchConfidence ? (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: cStyle.bg, color: cStyle.color }}>{item.matchConfidence}</span>
                    ) : item.templateApplied ? (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f0fdf4', color: '#059669' }}>Applied</span>
                    ) : (
                      <span style={{ fontSize: 10, color: '#cbd5e1' }}>—</span>
                    )}
                  </div>

                  {/* Actions */}
                  <div style={{ display: 'flex', gap: 2 }}>
                    {(item.confirmedTemplateId || item.matchedTemplateId) && !item.templateApplied && (
                      <IconButton
                        iconProps={{ iconName: 'Accept' }}
                        title="Apply template"
                        onClick={() => this.applyTemplateToItem(item.id, item.confirmedTemplateId || item.matchedTemplateId!)}
                        styles={{ root: { width: 26, height: 26 }, icon: { fontSize: 12, color: '#059669' } }}
                      />
                    )}
                    <IconButton
                      iconProps={{ iconName: 'Cancel' }}
                      title="Remove"
                      onClick={() => this.removeImport(item.id)}
                      styles={{ root: { width: 26, height: 26 }, icon: { fontSize: 12, color: '#dc2626' } }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {displayItems.length === 0 && !classifying && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 40, textAlign: 'center' }}>
            <Text style={{ color: '#94a3b8', fontSize: 13 }}>No classified items yet. Click "Classify" to start AI analysis.</Text>
          </div>
        )}
      </>
    );
  }

  // ---- REVIEW PHASE ----
  private renderReviewPhase(siteUrl: string): React.ReactElement {
    const { imports, selectedIds, searchQuery } = this.state;

    let filtered = imports;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(i => (i.policyTitle || i.fileName).toLowerCase().includes(q));
    }

    const allSelected = filtered.length > 0 && filtered.every(i => selectedIds.has(i.id));

    return (
      <>
        {/* Toolbar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
          <SearchBox placeholder="Search imports..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 220 } }} />
          <div style={{ flex: 1 }} />
          {selectedIds.size > 0 && (
            <>
              <span style={{ fontSize: 12, color: '#64748b', fontWeight: 600 }}>{selectedIds.size} selected</span>
              <DefaultButton text="Batch Assign Metadata" iconProps={{ iconName: 'Tag' }} onClick={() => this.setState({ showBatchPanel: true })} styles={{ root: { fontSize: 12, height: 30, borderRadius: 4 } }} />
            </>
          )}
          <PrimaryButton
            text="Open in Policy Builder"
            iconProps={{ iconName: 'EditCreate' }}
            disabled={selectedIds.size !== 1}
            href={selectedIds.size === 1 ? `${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${imports.find(i => selectedIds.has(i.id))?.spId || ''}` : undefined}
            styles={{ root: { fontSize: 12, height: 30, background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
          />
        </div>

        {/* Table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{
            display: 'grid', gridTemplateColumns: '36px 1fr 120px 100px 120px 90px 100px',
            padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
            fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b'
          }}>
            <div><input type="checkbox" checked={allSelected} onChange={() => {
              if (allSelected) this.setState({ selectedIds: new Set() });
              else this.setState({ selectedIds: new Set(filtered.map(i => i.id)) });
            }} /></div>
            <div>Policy</div>
            <div>Category</div>
            <div>Risk</div>
            <div>Departments</div>
            <div>Status</div>
            <div>Actions</div>
          </div>

          {filtered.map(item => {
            const isSelected = selectedIds.has(item.id);
            const statusColor = item.status === 'metadata-complete' ? '#059669' : item.status === 'classified' ? '#7c3aed' : item.status === 'failed' ? '#dc2626' : '#2563eb';
            const statusLabel = item.status === 'metadata-complete' ? 'Complete' : item.status === 'classified' ? 'Classified' : item.status === 'failed' ? 'Failed' : 'Uploaded';
            return (
              <div key={item.id} style={{
                display: 'grid', gridTemplateColumns: '36px 1fr 120px 100px 120px 90px 100px',
                padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center',
                background: isSelected ? '#f0fdfa' : '#fff'
              }}>
                <div><input type="checkbox" checked={isSelected} onChange={() => {
                  const next = new Set(selectedIds);
                  if (next.has(item.id)) next.delete(item.id); else next.add(item.id);
                  this.setState({ selectedIds: next });
                }} /></div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{item.policyTitle || item.fileName}</div>
                  <div style={{ fontSize: 11, color: '#94a3b8' }}>{item.fileName} • {item.fileType.toUpperCase()}</div>
                </div>
                <div style={{ fontSize: 12, color: '#475569' }}>{item.confirmedCategory || item.suggestedCategory || '-'}</div>
                <div style={{ fontSize: 12, color: '#475569' }}>{item.confirmedRisk || item.suggestedRisk || '-'}</div>
                <div style={{ fontSize: 11, color: '#64748b' }}>{(item.confirmedDepartments || item.suggestedDepartments || []).join(', ') || '-'}</div>
                <div><span style={{ fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: `${statusColor}15`, color: statusColor }}>{statusLabel}</span></div>
                <div style={{ display: 'flex', gap: 2 }}>
                  {item.spId && (
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Open in Policy Builder" href={`${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${item.spId}`}
                      styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#0d9488' } }} />
                  )}
                  {item.status === 'classified' && (
                    <IconButton iconProps={{ iconName: 'Accept' }} title="Accept AI suggestions" onClick={() => this.acceptAISuggestions(item.id)}
                      styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#059669' } }} />
                  )}
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" onClick={() => this.removeImport(item.id)}
                    styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#dc2626' } }} />
                </div>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // ---- BATCH PANEL ----
  private renderBatchPanel(): React.ReactElement {
    const { showBatchPanel, batchCategory, batchRisk, batchTemplateId, selectedIds, fastTrackTemplates } = this.state;

    const templateOptions: IDropdownOption[] = [
      { key: '', text: '— No template —' },
      ...fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title || t.ProfileName }))
    ];

    const hasTemplateSelected = !!batchTemplateId;
    const hasMetadata = !!batchCategory || !!batchRisk;

    return (
      <StyledPanel
        isOpen={showBatchPanel}
        onDismiss={() => this.setState({ showBatchPanel: false })}
        headerText={`Batch Assign (${selectedIds.size} selected)`}
        type={PanelType.smallFixedFar}
        onRenderFooterContent={() => (
          <Stack tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            {hasTemplateSelected && (
              <PrimaryButton
                text="Apply Template to Selected"
                iconProps={{ iconName: 'Tag' }}
                onClick={async () => {
                  await this.applyTemplateToSelected(parseInt(batchTemplateId));
                  this.setState({ showBatchPanel: false, batchTemplateId: '' });
                }}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488' } }}
              />
            )}
            {hasMetadata && (
              <DefaultButton
                text="Apply Metadata Only"
                onClick={() => { this.applyBatchMetadata(); }}
              />
            )}
            <DefaultButton text="Cancel" onClick={() => this.setState({ showBatchPanel: false })} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 20 }} style={{ paddingTop: 16 }}>
          <Text style={{ fontSize: 13, color: '#64748b' }}>
            Apply a Fast Track Template or set metadata for {selectedIds.size} selected polic{selectedIds.size !== 1 ? 'ies' : 'y'}.
          </Text>

          {/* Fast Track Template selector */}
          <div style={{ background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 8, padding: 16 }}>
            <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 13, display: 'block', marginBottom: 8 }}>Fast Track Template</Text>
            <Dropdown
              selectedKey={batchTemplateId}
              options={templateOptions}
              onChange={(_, opt) => this.setState({ batchTemplateId: String(opt?.key || '') })}
              placeholder="Select a template..."
            />
            <Text style={{ fontSize: 11, color: '#64748b', marginTop: 6, display: 'block' }}>
              Applies category, risk, read timeframe, acknowledgement, and quiz settings from the template.
            </Text>
          </div>

          <Text style={{ fontSize: 12, color: '#94a3b8', textAlign: 'center' }}>— or set individual fields —</Text>

          <Dropdown label="Category" selectedKey={batchCategory} options={CATEGORY_OPTIONS}
            onChange={(_, opt) => this.setState({ batchCategory: String(opt?.key || '') })} />
          <Dropdown label="Compliance Risk" selectedKey={batchRisk} options={RISK_OPTIONS}
            onChange={(_, opt) => this.setState({ batchRisk: String(opt?.key || '') })} />
        </Stack>
      </StyledPanel>
    );
  }

  // ---- HELPERS ----
  private getFileIcon(ext: string): string {
    switch (ext.toLowerCase()) {
      case '.docx': case '.doc': return 'WordDocument';
      case '.xlsx': case '.xls': return 'ExcelDocument';
      case '.pptx': case '.ppt': return 'PowerPointDocument';
      case '.pdf': return 'PDF';
      case '.txt': case '.rtf': return 'TextDocument';
      default: return 'Document';
    }
  }
}
