// @ts-nocheck
/**
 * PolicyRequestWizard — Extracted wizard component for requesting new policies.
 * Stepped form that persists to PM_PolicyRequests SharePoint list via PolicyRequestService.
 */
import * as React from 'react';
import styles from './PolicyManagerHeader.module.scss';
import { SPFI } from '@pnp/sp';
import { PolicyRequestService } from '../../services/PolicyRequestService';
import { IPolicyRequestFormData, IPolicyRequestSubmitResult, DEFAULT_REQUEST_FORM } from '../../models/IPolicyRequest';

export interface IPolicyRequestWizardProps {
  /** Whether the wizard overlay is visible */
  isOpen: boolean;
  /** Called when the wizard should close */
  onClose: () => void;
  /** PnPjs SPFI instance for SharePoint operations */
  sp?: SPFI;
  /** Current user's display name */
  userName?: string;
  /** Current user's email */
  userEmail?: string;
}

const DRAFT_STORAGE_KEY = 'pm_policy_request_draft';

const WIZARD_STEPS = [
  { title: 'Policy Details', description: 'What policy do you need?' },
  { title: 'Business Case', description: 'Why is this policy needed?' },
  { title: 'Requirements', description: 'Audience, timeline & compliance' },
  { title: 'Review & Submit', description: 'Confirm and submit your request' }
];

export const PolicyRequestWizard: React.FC<IPolicyRequestWizardProps> = ({
  isOpen,
  onClose,
  sp,
  userName = 'User',
  userEmail = ''
}) => {
  const [wizardStep, setWizardStep] = React.useState(0);
  const [wizardSubmitted, setWizardSubmitted] = React.useState(false);
  const [wizardSubmitting, setWizardSubmitting] = React.useState(false);
  const [wizardError, setWizardError] = React.useState<string | null>(null);
  const [submitResult, setSubmitResult] = React.useState<IPolicyRequestSubmitResult | null>(null);
  const [showUnsavedWarning, setShowUnsavedWarning] = React.useState(false);
  const [attachments, setAttachments] = React.useState<File[]>([]);
  const [isDragOver, setIsDragOver] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [requestForm, setRequestForm] = React.useState<IPolicyRequestFormData>(() => {
    try {
      const saved = localStorage.getItem(DRAFT_STORAGE_KEY);
      if (saved) return JSON.parse(saved);
    } catch { /* ignore */ }
    return { ...DEFAULT_REQUEST_FORM };
  });

  const updateRequestForm = (field: string, value: string | boolean) => {
    setRequestForm(prev => {
      const updated = { ...prev, [field]: value };
      try { localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(updated)); } catch { /* ignore */ }
      return updated;
    });
  };

  const formHasData = React.useMemo(() => {
    return requestForm.policyTitle.trim() !== '' ||
      requestForm.policyCategory !== '' ||
      requestForm.businessJustification.trim() !== '' ||
      requestForm.targetAudience.trim() !== '' ||
      requestForm.regulatoryDriver.trim() !== '' ||
      requestForm.additionalNotes.trim() !== '' ||
      attachments.length > 0;
  }, [requestForm, attachments]);

  const resetWizard = React.useCallback(() => {
    setWizardStep(0);
    setWizardSubmitted(false);
    setWizardSubmitting(false);
    setWizardError(null);
    setSubmitResult(null);
    setShowUnsavedWarning(false);
    setAttachments([]);
    setIsDragOver(false);
    setRequestForm({ ...DEFAULT_REQUEST_FORM });
    try { localStorage.removeItem(DRAFT_STORAGE_KEY); } catch { /* ignore */ }
  }, []);

  // Reset wizard state when opened
  React.useEffect(() => {
    if (isOpen) {
      setWizardStep(0);
      setWizardSubmitted(false);
      setWizardSubmitting(false);
      setWizardError(null);
      setSubmitResult(null);
      setShowUnsavedWarning(false);
      setAttachments([]);
      setIsDragOver(false);
      // Restore draft if available, otherwise reset
      try {
        const saved = localStorage.getItem(DRAFT_STORAGE_KEY);
        if (saved) {
          setRequestForm(JSON.parse(saved));
        } else {
          setRequestForm({ ...DEFAULT_REQUEST_FORM });
        }
      } catch {
        setRequestForm({ ...DEFAULT_REQUEST_FORM });
      }
    }
  }, [isOpen]);

  const handleWizardClose = () => {
    if (formHasData && !wizardSubmitted) {
      setShowUnsavedWarning(true);
    } else {
      resetWizard();
      onClose();
    }
  };

  const handleDiscardDraft = () => {
    setShowUnsavedWarning(false);
    resetWizard();
    onClose();
  };

  const getStepValidation = (step: number): { valid: boolean; errors: string[] } => {
    const errors: string[] = [];
    if (step === 0) {
      if (!requestForm.policyTitle.trim()) errors.push('Policy title is required');
      else if (requestForm.policyTitle.trim().length < 5) errors.push('Policy title must be at least 5 characters');
      if (!requestForm.policyCategory) errors.push('Policy category is required');
    }
    if (step === 1) {
      if (!requestForm.businessJustification.trim()) errors.push('Business justification is required');
      else if (requestForm.businessJustification.trim().length < 20) errors.push('Business justification must be at least 20 characters');
    }
    if (step === 2) {
      if (!requestForm.targetAudience.trim()) errors.push('Target audience is required');
      if (requestForm.desiredEffectiveDate) {
        const selectedDate = new Date(requestForm.desiredEffectiveDate);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (selectedDate < today) errors.push('Desired effective date must be in the future');
      }
    }
    return { valid: errors.length === 0, errors };
  };

  // Attachment helpers
  const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB per file
  const MAX_FILES = 5;
  const ALLOWED_TYPES = [
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'image/png',
    'image/jpeg',
    'text/plain'
  ];

  const addFiles = (files: FileList | File[]) => {
    const newFiles = Array.from(files).filter(file => {
      if (file.size > MAX_FILE_SIZE) return false;
      if (ALLOWED_TYPES.length > 0 && !ALLOWED_TYPES.includes(file.type)) return false;
      if (attachments.some(a => a.name === file.name && a.size === file.size)) return false;
      return true;
    });
    setAttachments(prev => [...prev, ...newFiles].slice(0, MAX_FILES));
  };

  const removeAttachment = (index: number) => {
    setAttachments(prev => prev.filter((_, i) => i !== index));
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const getFileIcon = (type: string): string => {
    if (type.includes('pdf')) return 'PDF';
    if (type.includes('word') || type.includes('document')) return 'DOC';
    if (type.includes('excel') || type.includes('sheet')) return 'XLS';
    if (type.includes('powerpoint') || type.includes('presentation')) return 'PPT';
    if (type.includes('image')) return 'IMG';
    return 'FILE';
  };

  const handleSubmitRequest = async () => {
    setWizardSubmitting(true);
    setWizardError(null);

    if (!sp) {
      const mockRef = `PR-${new Date().toISOString().slice(0,10).replace(/-/g,'')}-${Math.random().toString(36).substring(2, 7).toUpperCase()}`;
      setSubmitResult({ success: true, referenceNumber: mockRef, itemId: 0 });
      setWizardSubmitted(true);
      setWizardSubmitting(false);
      try { localStorage.removeItem(DRAFT_STORAGE_KEY); } catch { /* ignore */ }
      return;
    }

    try {
      const service = new PolicyRequestService(sp);
      const result = await service.submitRequest(requestForm, userName, userEmail, attachments.length > 0 ? attachments : undefined);

      if (result.success) {
        setSubmitResult(result);
        setWizardSubmitted(true);
        try { localStorage.removeItem(DRAFT_STORAGE_KEY); } catch { /* ignore */ }
      } else {
        setWizardError(result.error || 'Failed to submit request. Please try again.');
      }
    } catch (err) {
      setWizardError(err instanceof Error ? err.message : 'An unexpected error occurred. Please try again.');
    } finally {
      setWizardSubmitting(false);
    }
  };

  if (!isOpen) return null;

  return (
    <>
      {/* ================================================================ */}
      {/* REQUEST POLICY WIZARD — Full-screen overlay with stepped form     */}
      {/* ================================================================ */}
      <div className={styles.wizardOverlay} onClick={handleWizardClose}>
        <div className={styles.wizardModal} onClick={e => e.stopPropagation()}>
          {/* Wizard Header */}
          <div className={styles.wizardHeader}>
            <div className={styles.wizardHeaderLeft}>
              <div className={styles.wizardHeaderIcon}>
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 24, height: 24 }}>
                  <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  <path d="M14 2v6h6M12 18v-6M9 15h6" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <div>
                <div className={styles.wizardTitle}>Request a New Policy</div>
                <div className={styles.wizardSubtitle}>Submit a request to the Policy Authoring team</div>
              </div>
            </div>
            <button className={styles.wizardCloseBtn} onClick={handleWizardClose} title="Close">
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 20, height: 20 }}>
                <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </button>
          </div>

          {/* Step Progress */}
          {!wizardSubmitted && (
            <div className={styles.wizardStepper}>
              {WIZARD_STEPS.map((step, index) => (
                <div key={index} className={styles.wizardStepItem}>
                  <div className={`${styles.wizardStepCircle} ${index < wizardStep ? styles.wizardStepCompleted : ''} ${index === wizardStep ? styles.wizardStepActive : ''}`}>
                    {index < wizardStep ? (
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                        <path d="M20 6L9 17l-5-5" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    ) : (
                      <span>{index + 1}</span>
                    )}
                  </div>
                  <div className={styles.wizardStepLabel}>
                    <div className={styles.wizardStepTitle}>{step.title}</div>
                    <div className={styles.wizardStepDesc}>{step.description}</div>
                  </div>
                  {index < WIZARD_STEPS.length - 1 && <div className={`${styles.wizardStepConnector} ${index < wizardStep ? styles.wizardStepConnectorDone : ''}`} />}
                </div>
              ))}
            </div>
          )}

          {/* Wizard Body */}
          <div className={styles.wizardBody}>
            {wizardSubmitted ? (
              /* ===== SUCCESS STATE ===== */
              <div className={styles.wizardSuccess}>
                <div className={styles.wizardSuccessIcon}>
                  <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 48, height: 48 }}>
                    <path d="M22 11.08V12a10 10 0 11-5.93-9.14" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    <path d="M22 4L12 14.01l-3-3" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <h2 style={{ color: '#0f172a', margin: '16px 0 8px' }}>Policy Request Submitted!</h2>
                {submitResult?.referenceNumber && (
                  <div style={{ background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 8, padding: '8px 16px', display: 'inline-block', margin: '0 auto 16px', fontFamily: 'monospace', fontSize: 15, color: '#0d9488', fontWeight: 600, letterSpacing: 1 }}>
                    {submitResult.referenceNumber}
                  </div>
                )}
                <p style={{ color: '#64748b', maxWidth: 400, margin: '0 auto 24px', lineHeight: 1.6 }}>
                  Your request for "<strong>{requestForm.policyTitle}</strong>" has been submitted successfully.
                  The Policy Authoring team will be notified and will review your request shortly.
                </p>
                <div style={{ background: '#f0fdfa', borderRadius: 12, padding: 20, maxWidth: 420, margin: '0 auto 24px', textAlign: 'left' as const }}>
                  <div style={{ fontWeight: 600, marginBottom: 12, color: '#0d9488' }}>What happens next?</div>
                  <div style={{ display: 'flex', gap: 12, marginBottom: 10 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>1</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>A Policy Author will be assigned to your request</div>
                  </div>
                  <div style={{ display: 'flex', gap: 12, marginBottom: 10 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>2</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>They will draft the policy based on your requirements</div>
                  </div>
                  <div style={{ display: 'flex', gap: 12 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>3</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>You'll be notified when the draft is ready for review</div>
                  </div>
                </div>
                <button
                  className={styles.wizardBtnPrimary}
                  onClick={() => { resetWizard(); onClose(); }}
                >
                  Done
                </button>
              </div>
            ) : (
              <>
                {/* ===== STEP 0: Policy Details ===== */}
                {wizardStep === 0 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Policy Title <span className={styles.wizardRequired}>*</span></label>
                      <input
                        className={styles.wizardInput}
                        style={{ fontWeight: 400 }}
                        type="text"
                        placeholder="e.g. Data Retention Policy for Cloud Storage"
                        value={requestForm.policyTitle}
                        onChange={(e) => updateRequestForm('policyTitle', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Policy Category <span className={styles.wizardRequired}>*</span></label>
                      <select
                        className={styles.wizardSelect}
                        style={{ fontWeight: 400 }}
                        value={requestForm.policyCategory}
                        onChange={(e) => updateRequestForm('policyCategory', e.target.value)}
                      >
                        <option value="">Select category...</option>
                        <option value="IT Security">IT Security</option>
                        <option value="HR Policies">HR Policies</option>
                        <option value="Compliance">Compliance</option>
                        <option value="Health & Safety">Health & Safety</option>
                        <option value="Financial">Financial</option>
                        <option value="Legal">Legal</option>
                        <option value="Environmental">Environmental</option>
                        <option value="Operational">Operational</option>
                        <option value="Data Privacy">Data Privacy</option>
                        <option value="Quality Assurance">Quality Assurance</option>
                      </select>
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Request Type</label>
                      <select
                        className={styles.wizardSelect}
                        style={{ fontWeight: 400 }}
                        value={requestForm.policyType}
                        onChange={(e) => updateRequestForm('policyType', e.target.value)}
                      >
                        <option value="New Policy">New Policy</option>
                        <option value="Policy Update">Policy Update / Revision</option>
                        <option value="Policy Review">Policy Review</option>
                        <option value="Policy Replacement">Policy Replacement</option>
                      </select>
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Priority</label>
                      <select
                        className={styles.wizardSelect}
                        style={{ fontWeight: 400 }}
                        value={requestForm.priority}
                        onChange={(e) => updateRequestForm('priority', e.target.value)}
                      >
                        <option value="Low">Low — No urgency</option>
                        <option value="Medium">Medium — Standard timeline</option>
                        <option value="High">High — Urgent requirement</option>
                        <option value="Critical">Critical — Regulatory deadline</option>
                      </select>
                    </div>
                  </div>
                )}

                {/* ===== STEP 1: Business Case ===== */}
                {wizardStep === 1 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Business Justification <span className={styles.wizardRequired}>*</span></label>
                      <textarea
                        className={styles.wizardTextarea}
                        style={{ fontWeight: 400 }}
                        rows={5}
                        placeholder="Explain why this policy is needed. Include business drivers, risks of not having this policy, and any relevant context..."
                        value={requestForm.businessJustification}
                        onChange={(e) => updateRequestForm('businessJustification', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Regulatory / Compliance Driver</label>
                      <input
                        className={styles.wizardInput}
                        style={{ fontWeight: 400 }}
                        type="text"
                        placeholder="e.g. GDPR Article 5, ISO 27001, Health & Safety Act"
                        value={requestForm.regulatoryDriver}
                        onChange={(e) => updateRequestForm('regulatoryDriver', e.target.value)}
                      />
                      <div className={styles.wizardHelpText}>If this policy is driven by a regulatory requirement, specify the regulation or standard</div>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Additional Notes</label>
                      <textarea
                        className={styles.wizardTextarea}
                        style={{ fontWeight: 400 }}
                        rows={3}
                        placeholder="Any additional context, reference documents, or specific requirements for the policy author..."
                        value={requestForm.additionalNotes}
                        onChange={(e) => updateRequestForm('additionalNotes', e.target.value)}
                      />
                    </div>

                    {/* Attachment Upload */}
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>
                        Supporting Documents
                        <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 400, marginLeft: 8 }}>Optional — max {MAX_FILES} files, 10MB each</span>
                      </label>
                      <div
                        style={{
                          border: `2px dashed ${isDragOver ? '#0d9488' : '#e2e8f0'}`,
                          borderRadius: 10,
                          padding: attachments.length > 0 ? '12px 16px' : '24px 16px',
                          textAlign: 'center' as const,
                          background: isDragOver ? '#f0fdfa' : '#fafafa',
                          transition: 'all 0.2s ease',
                          cursor: attachments.length >= MAX_FILES ? 'default' : 'pointer'
                        }}
                        onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
                        onDragLeave={() => setIsDragOver(false)}
                        onDrop={(e) => {
                          e.preventDefault();
                          setIsDragOver(false);
                          if (e.dataTransfer.files.length > 0) addFiles(e.dataTransfer.files);
                        }}
                        onClick={() => {
                          if (attachments.length < MAX_FILES && fileInputRef.current) fileInputRef.current.click();
                        }}
                      >
                        <input
                          ref={fileInputRef}
                          type="file"
                          multiple
                          style={{ display: 'none' }}
                          accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.png,.jpg,.jpeg,.txt"
                          onChange={(e) => {
                            if (e.target.files) addFiles(e.target.files);
                            e.target.value = '';
                          }}
                        />
                        {attachments.length === 0 ? (
                          <div>
                            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 28, height: 28, margin: '0 auto 8px', display: 'block', color: '#94a3b8' }}>
                              <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                            </svg>
                            <div style={{ fontSize: 13, color: '#64748b' }}>
                              <span style={{ color: '#0d9488', fontWeight: 600 }}>Click to browse</span> or drag and drop files here
                            </div>
                            <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 4 }}>
                              PDF, Word, Excel, PowerPoint, images, or text files
                            </div>
                          </div>
                        ) : (
                          <div style={{ textAlign: 'left' as const }} onClick={e => e.stopPropagation()}>
                            {attachments.map((file, idx) => (
                              <div key={`${file.name}-${idx}`} style={{
                                display: 'flex', alignItems: 'center', gap: 10, padding: '6px 0',
                                borderBottom: idx < attachments.length - 1 ? '1px solid #f1f5f9' : 'none'
                              }}>
                                <div style={{
                                  width: 32, height: 32, borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center',
                                  fontSize: 10, fontWeight: 700, flexShrink: 0,
                                  background: file.type.includes('pdf') ? '#fef2f2' : file.type.includes('word') ? '#eff6ff' : file.type.includes('excel') || file.type.includes('sheet') ? '#f0fdf4' : file.type.includes('image') ? '#faf5ff' : '#f8fafc',
                                  color: file.type.includes('pdf') ? '#dc2626' : file.type.includes('word') ? '#2563eb' : file.type.includes('excel') || file.type.includes('sheet') ? '#16a34a' : file.type.includes('image') ? '#9333ea' : '#64748b'
                                }}>
                                  {getFileIcon(file.type)}
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: 13, color: '#334155', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{file.name}</div>
                                  <div style={{ fontSize: 11, color: '#94a3b8' }}>{formatFileSize(file.size)}</div>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => removeAttachment(idx)}
                                  style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4, color: '#94a3b8', borderRadius: 4 }}
                                  title="Remove file"
                                >
                                  <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                                    <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                                  </svg>
                                </button>
                              </div>
                            ))}
                            {attachments.length < MAX_FILES && (
                              <button
                                type="button"
                                onClick={() => fileInputRef.current?.click()}
                                style={{
                                  background: 'none', border: '1px dashed #cbd5e1', borderRadius: 6, padding: '6px 12px',
                                  fontSize: 12, color: '#0d9488', cursor: 'pointer', marginTop: 8, width: '100%'
                                }}
                              >
                                + Add more files ({MAX_FILES - attachments.length} remaining)
                              </button>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                )}

                {/* ===== STEP 2: Requirements ===== */}
                {wizardStep === 2 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Target Audience <span className={styles.wizardRequired}>*</span></label>
                      <input
                        className={styles.wizardInput}
                        style={{ fontWeight: 400 }}
                        type="text"
                        placeholder="e.g. All Employees, IT Department, Management, Contractors"
                        value={requestForm.targetAudience}
                        onChange={(e) => updateRequestForm('targetAudience', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Desired Effective Date</label>
                      <input
                        className={styles.wizardInput}
                        style={{ fontWeight: 400 }}
                        type="date"
                        value={requestForm.desiredEffectiveDate}
                        onChange={(e) => updateRequestForm('desiredEffectiveDate', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Read Timeframe (days)</label>
                      <select
                        className={styles.wizardSelect}
                        style={{ fontWeight: 400 }}
                        value={requestForm.readTimeframeDays}
                        onChange={(e) => updateRequestForm('readTimeframeDays', e.target.value)}
                      >
                        <option value="7">7 days</option>
                        <option value="14">14 days</option>
                        <option value="30">30 days</option>
                        <option value="60">60 days</option>
                        <option value="90">90 days</option>
                      </select>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <div className={styles.wizardCheckboxGroup}>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.requiresAcknowledgement}
                            onChange={(e) => updateRequestForm('requiresAcknowledgement', e.target.checked)}
                          />
                          <span>Require employee acknowledgement</span>
                        </label>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.requiresQuiz}
                            onChange={(e) => updateRequestForm('requiresQuiz', e.target.checked)}
                          />
                          <span>Require comprehension quiz</span>
                        </label>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.notifyAuthors}
                            onChange={(e) => updateRequestForm('notifyAuthors', e.target.checked)}
                          />
                          <span>Notify Policy Authors immediately</span>
                        </label>
                      </div>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel} style={{ fontWeight: 400 }}>Preferred Author (optional)</label>
                      <input
                        className={styles.wizardInput}
                        style={{ fontWeight: 400 }}
                        type="text"
                        placeholder="Leave blank to auto-assign, or enter a name"
                        value={requestForm.preferredAuthor}
                        onChange={(e) => updateRequestForm('preferredAuthor', e.target.value)}
                      />
                    </div>
                  </div>
                )}

                {/* ===== STEP 3: Review & Submit ===== */}
                {wizardStep === 3 && (
                  <div className={styles.wizardReview}>
                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M14 2v6h6" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        Policy Details
                      </div>
                      <div className={styles.wizardReviewGrid}>
                        <div><span className={styles.wizardReviewLabel}>Title:</span> <strong>{requestForm.policyTitle}</strong></div>
                        <div><span className={styles.wizardReviewLabel}>Category:</span> {requestForm.policyCategory}</div>
                        <div><span className={styles.wizardReviewLabel}>Type:</span> {requestForm.policyType}</div>
                        <div><span className={styles.wizardReviewLabel}>Priority:</span>
                          <span style={{
                            padding: '2px 10px', borderRadius: 10, fontSize: 12, fontWeight: 600, marginLeft: 6,
                            background: requestForm.priority === 'Critical' ? '#fde7e9' : requestForm.priority === 'High' ? '#fff3e0' : requestForm.priority === 'Medium' ? '#fff8e1' : '#f1f5f9',
                            color: requestForm.priority === 'Critical' ? '#d13438' : requestForm.priority === 'High' ? '#f97316' : requestForm.priority === 'Medium' ? '#f59e0b' : '#64748b'
                          }}>
                            {requestForm.priority}
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        Business Case
                      </div>
                      <div style={{ fontSize: 13, lineHeight: 1.6, color: '#334155' }}>{requestForm.businessJustification}</div>
                      {requestForm.regulatoryDriver && (
                        <div style={{ marginTop: 8, fontSize: 12, color: '#ef4444' }}>
                          <strong>Regulatory Driver:</strong> {requestForm.regulatoryDriver}
                        </div>
                      )}
                      {requestForm.additionalNotes && (
                        <div style={{ marginTop: 8, fontSize: 12, color: '#64748b', fontStyle: 'italic' }}>
                          <strong>Notes:</strong> {requestForm.additionalNotes}
                        </div>
                      )}
                    </div>

                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <circle cx="9" cy="7" r="4" stroke="#0d9488" strokeWidth="2"/>
                        </svg>
                        Requirements
                      </div>
                      <div className={styles.wizardReviewGrid}>
                        <div><span className={styles.wizardReviewLabel}>Audience:</span> {requestForm.targetAudience}</div>
                        <div><span className={styles.wizardReviewLabel}>Effective Date:</span> {requestForm.desiredEffectiveDate ? new Date(requestForm.desiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : 'Not specified'}</div>
                        <div><span className={styles.wizardReviewLabel}>Read Timeframe:</span> {requestForm.readTimeframeDays} days</div>
                        <div><span className={styles.wizardReviewLabel}>Acknowledgement:</span> {requestForm.requiresAcknowledgement ? 'Yes' : 'No'}</div>
                        <div><span className={styles.wizardReviewLabel}>Quiz:</span> {requestForm.requiresQuiz ? 'Yes' : 'No'}</div>
                        {requestForm.preferredAuthor && <div><span className={styles.wizardReviewLabel}>Preferred Author:</span> {requestForm.preferredAuthor}</div>}
                      </div>
                    </div>

                    {attachments.length > 0 && (
                      <div className={styles.wizardReviewSection}>
                        <div className={styles.wizardReviewTitle}>
                          <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                            <path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          </svg>
                          Attachments ({attachments.length})
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column' as const, gap: 4 }}>
                          {attachments.map((file, idx) => (
                            <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: '#334155' }}>
                              <span style={{
                                fontSize: 10, fontWeight: 700, padding: '2px 6px', borderRadius: 4,
                                background: file.type.includes('pdf') ? '#fef2f2' : file.type.includes('word') ? '#eff6ff' : '#f8fafc',
                                color: file.type.includes('pdf') ? '#dc2626' : file.type.includes('word') ? '#2563eb' : '#64748b'
                              }}>{getFileIcon(file.type)}</span>
                              {file.name}
                              <span style={{ fontSize: 11, color: '#94a3b8' }}>({formatFileSize(file.size)})</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    <div style={{ background: '#fffbeb', borderRadius: 8, padding: 12, border: '1px solid #fde68a', display: 'flex', gap: 10, alignItems: 'flex-start', marginTop: 8 }}>
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 20, height: 20, flexShrink: 0, marginTop: 1 }}>
                        <path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      <div style={{ fontSize: 12, color: '#92400e', lineHeight: 1.5 }}>
                        <strong>Workflow notification:</strong> Upon submission, the Policy Authoring team will receive an email notification with your request details. You will be notified when an author is assigned and when the draft is ready for review.
                      </div>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>

          {/* Error Banner */}
          {wizardError && !wizardSubmitted && (
            <div style={{ background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 8, padding: '10px 16px', margin: '0 24px', display: 'flex', gap: 10, alignItems: 'center' }}>
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 20, height: 20, flexShrink: 0 }}>
                <path d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" stroke="#ef4444" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              <div style={{ fontSize: 13, color: '#991b1b', flex: 1 }}>{wizardError}</div>
              <button onClick={() => setWizardError(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#991b1b', padding: 4 }}>
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                  <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </button>
            </div>
          )}

          {/* Wizard Footer */}
          {!wizardSubmitted && (
            <div className={styles.wizardFooter}>
              <div>
                {wizardStep > 0 && (
                  <button className={styles.wizardBtnSecondary} onClick={() => setWizardStep(wizardStep - 1)} disabled={wizardSubmitting}>
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                      <path d="M19 12H5M12 19l-7-7 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                    Back
                  </button>
                )}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className={styles.wizardBtnOutline} onClick={handleWizardClose} disabled={wizardSubmitting}>
                  Cancel
                </button>
                {wizardStep < WIZARD_STEPS.length - 1 ? (
                  <button
                    className={styles.wizardBtnPrimary}
                    onClick={() => setWizardStep(wizardStep + 1)}
                    disabled={!getStepValidation(wizardStep).valid}
                  >
                    Next
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                      <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  </button>
                ) : (
                  <button
                    className={styles.wizardBtnSubmit}
                    onClick={handleSubmitRequest}
                    disabled={wizardSubmitting}
                  >
                    {wizardSubmitting ? (
                      <>
                        <span style={{ width: 14, height: 14, border: '2px solid rgba(255,255,255,0.3)', borderTopColor: '#fff', borderRadius: '50%', display: 'inline-block', animation: 'spin 1s linear infinite' }} />
                        Submitting...
                      </>
                    ) : (
                      <>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                          <path d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        Submit Request
                      </>
                    )}
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Unsaved Changes Confirmation Dialog */}
      {showUnsavedWarning && (
        <div className={styles.wizardOverlay} style={{ zIndex: 10002 }} onClick={() => setShowUnsavedWarning(false)}>
          <div style={{ background: '#fff', borderRadius: 16, padding: 32, maxWidth: 420, width: '90%', textAlign: 'center' as const, boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)' }} onClick={e => e.stopPropagation()}>
            <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px' }}>
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 24, height: 24 }}>
                <path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </div>
            <h3 style={{ color: '#0f172a', margin: '0 0 8px', fontSize: 18 }}>Unsaved Changes</h3>
            <p style={{ color: '#64748b', margin: '0 0 24px', fontSize: 14, lineHeight: 1.6 }}>
              You have unsaved changes in your policy request. Do you want to discard them or continue editing?
            </p>
            <div style={{ display: 'flex', gap: 12, justifyContent: 'center' }}>
              <button
                className={styles.wizardBtnOutline}
                onClick={handleDiscardDraft}
                style={{ flex: 1 }}
              >
                Discard
              </button>
              <button
                className={styles.wizardBtnPrimary}
                onClick={() => setShowUnsavedWarning(false)}
                style={{ flex: 1 }}
              >
                Continue Editing
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default PolicyRequestWizard;
