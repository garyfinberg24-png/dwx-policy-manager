// @ts-nocheck
import * as React from 'react';
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
  DefaultButton, PrimaryButton, IconButton,
  TextField, Dropdown, IDropdownOption, Checkbox, Label, Icon, Toggle
} from '@fluentui/react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  ComplianceRisk, ReadTimeframe, PolicyCategory, PolicyStatus
} from '../../../../models/IPolicy';
import {
  WIZARD_STEPS, ICorporateTemplate,
} from '../../../../models/IPolicyAuthor';
import { PEOPLE_PICKER } from '../../../../constants/PolicyAuthorConstants';

// Step field descriptions for accordion sidebar
const STEP_FIELDS: Record<number, string[]> = {
  0: ['Choose creation method', 'Template, blank, upload, or Office document'],
  1: ['Policy number', 'Name & category', 'Summary', 'Risk level'],
  2: ['Rich text editor', 'Key points', 'Supporting documents'],
  3: ['Compliance risk', 'Read timeframe', 'Acknowledgement', 'Quiz'],
  4: ['Target all or specific groups', 'Departments & roles', 'Contractors'],
  5: ['Effective date', 'Expiry date', 'Review frequency', 'Supersedes'],
  6: ['Reviewers', 'Approvers'],
  7: ['Final review', 'Submit or save draft'],
};

export interface IPolicyWizardProps {
  // State values
  currentStep: number;
  completedSteps: Set<number>;
  creationMethod: string;
  creatingDocument: boolean;
  policyNumber: string;
  policyName: string;
  policyCategory: string;
  policySummary: string;
  policyContent: string;
  complianceRisk: string;
  readTimeframe: string;
  readTimeframeDays: number;
  requiresAcknowledgement: boolean;
  requiresQuiz: boolean;
  availableQuizzes: any[];
  availableQuizzesLoading: boolean;
  selectedQuizId: number | null;
  selectedQuizTitle: string;
  targetAllEmployees: boolean;
  targetDepartments: string[];
  targetRoles: string[];
  targetLocations: string[];
  includeContractors: boolean;
  effectiveDate: string;
  expiryDate: string;
  reviewFrequency: string;
  nextReviewDate: string;
  supersedesPolicy: string;
  browsePolicies: any[];
  reviewers: string[];
  approvers: string[];
  keyPoints: string[];
  newKeyPoint: string;
  linkedDocumentUrl: string;
  linkedDocumentType: string;
  expandedReviewSections: Set<string>;
  // Style module
  styles: Record<string, string>;
  // Context for PeoplePicker
  context: any;
  siteUrl: string;
  // Callbacks
  onSetState: (update: Record<string, any>) => void;
  onGoToStep: (step: number) => void;
  onNextStep: () => void;
  onPreviousStep: () => void;
  onSelectCreationMethod: (method: string) => void;
  onAddKeyPoint: () => void;
  onRemoveKeyPoint: (index: number) => void;
  onLoadAvailableQuizzes: () => void;
}

const PolicyWizard: React.FC<IPolicyWizardProps> = (props) => {
  const s = props.styles;

  // ──── Sidebar ────
  const renderSidebar = (): JSX.Element => {
    return (
      <aside className={s.v3Sidebar}>
        <div className={s.v3SidebarHeader}>
          <Text variant="mediumPlus" style={{ fontWeight: 700, color: '#111827', display: 'block' }}>New Policy Wizard</Text>
          <Text variant="small" style={{ color: '#6b7280', marginTop: 2 }}>{WIZARD_STEPS.length} steps to complete</Text>
        </div>
        <div className={s.v3Accordion}>
          {WIZARD_STEPS.map((step, index) => {
            const isCompleted = props.completedSteps.has(index);
            const isCurrent = index === props.currentStep;
            const isClickable = index <= props.currentStep || props.completedSteps.has(index - 1) || index === 0;
            const stateClass = isCompleted ? 'completed' : isCurrent ? 'active' : 'future';

            return (
              <div key={step.key} className={`${s.v3AccItem} ${s[`v3AccItem_${stateClass}`] || ''}`}>
                <div
                  className={s.v3AccHeader}
                  onClick={() => isClickable && props.onGoToStep(index)}
                  style={{ cursor: isClickable ? 'pointer' : 'default' }}
                >
                  <div
                    className={s.v3AccNum}
                    style={{
                      background: isCompleted || isCurrent ? '#0d9488' : '#e5e7eb',
                      color: isCompleted || isCurrent ? '#ffffff' : '#6b7280'
                    }}
                  >
                    {isCompleted ? <Icon iconName="CheckMark" style={{ fontSize: 11 }} /> : <span>{index + 1}</span>}
                  </div>
                  <span style={{
                    fontWeight: isCurrent ? 600 : 500,
                    color: isCompleted ? '#6b7280' : isCurrent ? '#0f766e' : '#374151',
                    fontSize: 13, flex: 1
                  }}>
                    {step.title}
                  </span>
                  <span style={{
                    fontSize: 10, color: '#9ca3af',
                    transition: 'transform 0.2s',
                    transform: isCurrent ? 'rotate(180deg)' : 'rotate(0deg)'
                  }}>&#9660;</span>
                </div>
                {isCurrent && STEP_FIELDS[index] && (
                  <div className={s.v3AccBody}>
                    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                      {STEP_FIELDS[index].map((field, fi) => (
                        <li key={fi} style={{
                          padding: '4px 0', fontSize: 12,
                          color: fi === 0 ? '#0f766e' : '#6b7280',
                          fontWeight: fi === 0 ? 600 : 400,
                          display: 'flex', alignItems: 'center', gap: 6
                        }}>
                          <span style={{
                            width: 5, height: 5, borderRadius: '50%',
                            background: '#0d9488', opacity: fi === 0 ? 1 : 0.5,
                            display: 'inline-block', flexShrink: 0
                          }} />
                          {field}
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </aside>
    );
  };

  // ──── Context Panel ────
  const renderContextPanel = (): JSX.Element => {
    const tipsMap: Record<number, { title: string; body: string }[]> = {
      0: [
        { title: 'Choosing a Method', body: 'Start from a template for consistency, or choose blank for full creative control.' },
        { title: 'Corporate Templates', body: 'Corporate templates include pre-approved branding, headers, and formatting.' }
      ],
      1: [
        { title: 'Policy Title Best Practices', body: 'Use descriptive, action-oriented titles. Avoid acronyms unless universally understood.' },
        { title: 'Category Selection', body: 'Choose the primary category that best represents the policy scope.' },
        { title: 'Writing a Good Summary', body: 'Include the policy\'s purpose, who it applies to, and the key actions. Aim for 2-3 sentences.' }
      ],
      2: [
        { title: 'Content Structure', body: 'Use clear headings and bullet points. Start with purpose, then scope, responsibilities, and procedures.' },
        { title: 'Key Points', body: 'Add 3-5 key points that summarize the most important takeaways for readers.' }
      ],
      3: [
        { title: 'Risk Assessment', body: 'Consider regulatory, legal, and operational risk. Higher risk = stricter compliance tracking.' },
        { title: 'Acknowledgement & Quiz', body: 'Critical policies should require both acknowledgement and quiz completion.' }
      ],
      4: [
        { title: 'Target Audience', body: 'Select "All Employees" for company-wide policies. For department-specific, choose relevant teams.' },
        { title: 'Contractors', body: 'If your policy applies to external contractors, include them in the audience.' }
      ],
      5: [
        { title: 'Effective Dates', body: 'Allow at least 2 weeks between publication and effective date.' },
        { title: 'Review Cycle', body: 'Most policies should be reviewed annually. Critical ones may need quarterly review.' }
      ],
      6: [
        { title: 'Review Workflow', body: 'Add subject matter experts as reviewers and department heads as approvers.' },
        { title: 'Multi-Level Approval', body: 'High-risk policies typically require both department and executive approval.' }
      ],
      7: [
        { title: 'Final Check', body: 'Review all sections carefully. Once submitted, the policy enters the review workflow.' },
        { title: 'Draft Option', body: 'Not ready to submit? Save as draft to continue editing later.' }
      ]
    };

    const tips = tipsMap[props.currentStep] || [];

    return (
      <aside className={s.v3RightPanel}>
        <div className={s.v3PanelSection}>
          <Text variant="small" style={{ fontWeight: 700, color: '#1f2937', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: '#0d9488', fontSize: 10
            }}>
              <Icon iconName="Lightbulb" style={{ fontSize: 12 }} />
            </span>
            Tips & Guidance
          </Text>
          {tips.map((tip, i) => (
            <div key={i} className={s.v3Tip}>
              <Text style={{ display: 'block', marginBottom: 4, fontSize: 12, fontWeight: 600 }}>{tip.title}</Text>
              <Text style={{ fontSize: 12, color: '#115e59', lineHeight: '1.5' }}>{tip.body}</Text>
            </div>
          ))}
        </div>
      </aside>
    );
  };

  // ──── Step Renderers ────

  const renderStep0 = (): JSX.Element => {
    const primaryMethods = [
      { key: 'blank', title: 'Blank Policy', description: 'Start with empty rich text editor', icon: 'Page', iconClass: 'iconBlank' },
      { key: 'template', title: 'From Template', description: 'Use a pre-approved policy template', icon: 'DocumentSet', iconClass: 'iconTemplate' },
      { key: 'upload', title: 'Upload Document', description: 'Import from Word, PDF, or other file', icon: 'Upload', iconClass: 'iconUpload' }
    ];
    const officeMethods = [
      { key: 'word', title: 'Word Document', description: 'Create new Word document', icon: 'WordDocument', iconClass: 'iconWord', color: '#2b579a' },
      { key: 'excel', title: 'Excel Spreadsheet', description: 'Create Excel for data policies', icon: 'ExcelDocument', iconClass: 'iconExcel', color: '#217346' },
      { key: 'powerpoint', title: 'PowerPoint', description: 'Create presentation-style policy', icon: 'PowerPointDocument', iconClass: 'iconPowerPoint', color: '#b7472a' }
    ];
    const additionalMethods = [
      { key: 'corporate', title: 'Corporate Template', description: 'Use branded company template', icon: 'FileTemplate', iconClass: 'iconCorporate' },
      { key: 'infographic', title: 'Infographic/Image', description: 'Visual policy (floor plans, etc.)', icon: 'PictureFill', iconClass: 'iconImage' }
    ];

    const getStyle = (name: string): string => (s as any)[name] || '';

    const renderMethodCard = (method: any) => (
      <div
        key={method.key}
        className={`${s.creationMethodCard} ${props.creationMethod === method.key ? s.selected : ''}`}
        onClick={() => props.onSelectCreationMethod(method.key)}
      >
        <div className={getStyle('creationMethodCardHeader')}>
          <div className={`${s.creationMethodIcon} ${getStyle(method.iconClass)}`}>
            <Icon iconName={method.icon} style={{ color: method.color || '#0078d4' }} />
          </div>
          <Text className={s.creationMethodTitle}>{method.title}</Text>
        </div>
        <Text className={s.creationMethodDescription}>{method.description}</Text>
      </div>
    );

    return (
      <div className={s.wizardStepContent}>
        {props.creatingDocument && (
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Creating document..." />
          </Stack>
        )}
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}><h3>Create New Policy</h3><span>Choose how to start</span></div>
          <div className={s.creationMethodGrid}>{primaryMethods.map(renderMethodCard)}</div>
        </div>
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}><h3>Create from Office</h3><span>For complex documents</span></div>
          <div className={s.creationMethodGrid}>{officeMethods.map(renderMethodCard)}</div>
        </div>
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}><h3>Additional Options</h3></div>
          <div className={s.creationMethodGrid}>{additionalMethods.map(renderMethodCard)}</div>
        </div>
        <div className={getStyle('methodTip')}>
          <Icon iconName="Info" />
          <span><strong>Tip:</strong> For most policies, "Blank Policy" or "From Template" are recommended.</span>
        </div>
      </div>
    );
  };

  const renderStep3 = (): JSX.Element => {
    const riskOptions: IDropdownOption[] = Object.values(ComplianceRisk).map(risk => ({ key: risk, text: risk }));
    const timeframeOptions: IDropdownOption[] = Object.values(ReadTimeframe).map(tf => ({ key: tf, text: tf }));

    return (
      <div className={s.wizardStepContent}>
        <div className={s.section}>
          <Stack tokens={{ childrenGap: 20 }}>
            <Dropdown label="Compliance Risk Level" required selectedKey={props.complianceRisk} options={riskOptions}
              onChange={(e, option) => props.onSetState({ complianceRisk: option?.key as string })} styles={{ root: { maxWidth: 300 } }} />
            <Dropdown label="Read Timeframe" selectedKey={props.readTimeframe} options={timeframeOptions}
              onChange={(e, option) => {
                const selected = option?.key as string;
                props.onSetState({ readTimeframe: selected, readTimeframeDays: selected === ReadTimeframe.Custom ? props.readTimeframeDays : 7 });
              }} styles={{ root: { maxWidth: 300 } }} />
            {props.readTimeframe === ReadTimeframe.Custom && (
              <TextField label="Custom Days" type="number" value={props.readTimeframeDays.toString()}
                onChange={(e, value) => props.onSetState({ readTimeframeDays: parseInt(value || '7', 10) })} styles={{ root: { maxWidth: 150 } }} />
            )}
            <Stack tokens={{ childrenGap: 12 }}>
              <Checkbox label="Requires Acknowledgement" checked={props.requiresAcknowledgement}
                onChange={(e, checked) => props.onSetState({ requiresAcknowledgement: checked || false })} />
              <Text variant="small" style={{ marginLeft: 26, color: '#605e5c' }}>Employees must confirm they have read and understood the policy</Text>
              <Checkbox label="Requires Quiz" checked={props.requiresQuiz}
                onChange={(e, checked) => {
                  props.onSetState({ requiresQuiz: checked || false });
                  if (checked) props.onLoadAvailableQuizzes();
                }} />
              <Text variant="small" style={{ marginLeft: 26, color: '#605e5c' }}>Employees must pass a quiz to demonstrate understanding</Text>
              {props.requiresQuiz && (
                <div style={{ marginLeft: 26, marginTop: 8, padding: 16, background: '#f0fdfa', borderRadius: 8, border: '1px solid #e2e8f0' }}>
                  <Text variant="medium" style={{ fontWeight: 600, color: '#0f172a', display: 'block', marginBottom: 12 }}>Quiz Selection</Text>
                  {props.availableQuizzesLoading ? (
                    <Spinner size={SpinnerSize.small} label="Loading quizzes..." />
                  ) : props.availableQuizzes.length > 0 ? (
                    <>
                      <Dropdown label="Select an existing quiz" placeholder="Choose a quiz..."
                        selectedKey={props.selectedQuizId ?? undefined}
                        options={[
                          { key: '', text: '— No quiz selected —' },
                          ...props.availableQuizzes.map(q => ({ key: q.Id, text: `${q.Title} (${q.QuestionCount} questions, ${q.PassingScore}% to pass)` }))
                        ]}
                        onChange={(_e, option) => {
                          if (option && option.key !== '') {
                            props.onSetState({
                              selectedQuizId: option.key as number,
                              selectedQuizTitle: props.availableQuizzes.find(q => q.Id === option.key)?.Title || ''
                            });
                          } else {
                            props.onSetState({ selectedQuizId: null, selectedQuizTitle: '' });
                          }
                        }}
                        styles={{ root: { maxWidth: 450, marginBottom: 12 } }} />
                      {props.selectedQuizId && (
                        <MessageBar messageBarType={MessageBarType.success} styles={{ root: { borderRadius: 4 } }}>
                          Quiz "{props.selectedQuizTitle}" will be assigned to this policy.
                        </MessageBar>
                      )}
                    </>
                  ) : (
                    <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 4, marginBottom: 12 } }}>
                      No published quizzes found. You can create one after the policy is published.
                    </MessageBar>
                  )}
                  <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
                    <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh Quizzes"
                      onClick={() => props.onLoadAvailableQuizzes()} styles={{ root: { height: 32 }, label: { fontSize: 12 } }} />
                    <DefaultButton iconProps={{ iconName: 'Add' }} text="Create New Quiz"
                      onClick={() => window.open(`${props.siteUrl}/SitePages/QuizBuilder.aspx`, '_blank')}
                      styles={{ root: { height: 32 }, label: { fontSize: 12 } }} />
                  </Stack>
                </div>
              )}
            </Stack>
          </Stack>
        </div>
      </div>
    );
  };

  const renderStep4 = (): JSX.Element => (
    <div className={s.wizardStepContent}>
      <div className={s.section}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Checkbox label="All Employees" checked={props.targetAllEmployees}
            onChange={(e, checked) => props.onSetState({ targetAllEmployees: checked || false })} />
          {!props.targetAllEmployees && (
            <>
              <TextField label="Target Departments" placeholder="e.g., HR, IT, Finance (comma-separated)"
                value={props.targetDepartments.join(', ')}
                onChange={(e, value) => props.onSetState({ targetDepartments: value ? value.split(',').map(d => d.trim()) : [] })} />
              <TextField label="Target Roles" placeholder="e.g., Manager, Director (comma-separated)"
                value={props.targetRoles.join(', ')}
                onChange={(e, value) => props.onSetState({ targetRoles: value ? value.split(',').map(r => r.trim()) : [] })} />
              <TextField label="Target Locations" placeholder="e.g., London, New York (comma-separated)"
                value={props.targetLocations.join(', ')}
                onChange={(e, value) => props.onSetState({ targetLocations: value ? value.split(',').map(l => l.trim()) : [] })} />
            </>
          )}
          <Checkbox label="Include Contractors/Third Parties" checked={props.includeContractors}
            onChange={(e, checked) => props.onSetState({ includeContractors: checked || false })} />
        </Stack>
      </div>
    </div>
  );

  const calcNextReviewDate = (effective: string, frequency: string): string => {
    if (!effective || frequency === 'None' || !frequency) return '';
    const date = new Date(effective);
    if (isNaN(date.getTime())) return '';
    switch (frequency) {
      case 'Annual': date.setMonth(date.getMonth() + 12); break;
      case 'Biannual': date.setMonth(date.getMonth() + 6); break;
      case 'Quarterly': date.setMonth(date.getMonth() + 3); break;
      case 'Monthly': date.setMonth(date.getMonth() + 1); break;
      default: return '';
    }
    return date.toISOString().split('T')[0];
  };

  const renderStep5 = (): JSX.Element => {
    const frequencyOptions: IDropdownOption[] = [
      { key: 'Annual', text: 'Annual (every 12 months)' },
      { key: 'Biannual', text: 'Biannual (every 6 months)' },
      { key: 'Quarterly', text: 'Quarterly (every 3 months)' },
      { key: 'Monthly', text: 'Monthly' },
      { key: 'None', text: 'No scheduled review' }
    ];
    const supersedesOptions: IDropdownOption[] = [
      { key: '', text: '(None)' },
      ...props.browsePolicies
        .filter(p => p.PolicyStatus === PolicyStatus.Published || p.PolicyStatus === PolicyStatus.Approved)
        .map(p => ({ key: p.PolicyNumber || p.Title, text: `${p.PolicyNumber || 'N/A'} — ${p.Title}` }))
    ];

    return (
      <div className={s.wizardStepContent}>
        <div className={s.section}>
          <Stack tokens={{ childrenGap: 20 }}>
            <TextField label="Effective Date" type="date" required value={props.effectiveDate}
              onChange={(e, value) => {
                const newEff = value || '';
                const computed = calcNextReviewDate(newEff, props.reviewFrequency);
                props.onSetState({ effectiveDate: newEff, nextReviewDate: computed || props.nextReviewDate });
              }} styles={{ root: { maxWidth: 200 } }} />
            <TextField label="Expiry Date (Optional)" type="date" value={props.expiryDate}
              onChange={(e, value) => props.onSetState({ expiryDate: value || '' })} styles={{ root: { maxWidth: 200 } }} />
            <Dropdown label="Review Frequency" selectedKey={props.reviewFrequency} options={frequencyOptions}
              onChange={(e, option) => {
                const freq = option?.key as string;
                const computed = calcNextReviewDate(props.effectiveDate, freq);
                props.onSetState({ reviewFrequency: freq, nextReviewDate: computed || props.nextReviewDate });
              }} styles={{ root: { maxWidth: 300 } }} />
            <TextField label="Next Review Date" type="date" value={props.nextReviewDate}
              onChange={(e, value) => props.onSetState({ nextReviewDate: value || '' })}
              description={props.effectiveDate && props.reviewFrequency && props.reviewFrequency !== 'None'
                ? `Auto-calculated from effective date + ${props.reviewFrequency.toLowerCase()} frequency.` : undefined}
              styles={{ root: { maxWidth: 200 } }} />
            <Dropdown label="Supersedes Policy (Optional)" placeholder="Select a policy this replaces..."
              selectedKey={props.supersedesPolicy || ''} options={supersedesOptions}
              onChange={(e, option) => props.onSetState({ supersedesPolicy: (option?.key as string) || '' })}
              styles={{ root: { maxWidth: 400 } }} />
          </Stack>
        </div>
      </div>
    );
  };

  const renderStep7 = (): JSX.Element => {
    const toggleSection = (key: string): void => {
      const next = new Set(props.expandedReviewSections);
      if (next.has(key)) next.delete(key); else next.add(key);
      props.onSetState({ expandedReviewSections: next });
    };
    const sections = [
      { key: 'basic', icon: 'Info', title: 'Basic Information', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Policy Number</Label><Text>{props.policyNumber || '(Auto-generated)'}</Text></div>
          <div className={s.reviewItem}><Label>Policy Name</Label><Text>{props.policyName || '-'}</Text></div>
          <div className={s.reviewItem}><Label>Category</Label><Text>{props.policyCategory || '-'}</Text></div>
          <div className={s.reviewItem}><Label>Summary</Label><Text>{props.policySummary || '-'}</Text></div>
        </div>
      )},
      { key: 'content', icon: 'Edit', title: 'Content', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Content Preview</Label><Text>{props.policyContent ? `${props.policyContent.substring(0, 200).replace(/<[^>]*>/g, '')}...` : '-'}</Text></div>
          {props.linkedDocumentUrl && <div className={s.reviewItem}><Label>Linked Document</Label><Text>{props.linkedDocumentType}: {props.linkedDocumentUrl}</Text></div>}
          <div className={s.reviewItem}><Label>Key Points</Label><Text>{props.keyPoints.length > 0 ? props.keyPoints.join(', ') : 'None specified'}</Text></div>
        </div>
      )},
      { key: 'compliance', icon: 'Shield', title: 'Compliance & Risk', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Risk Level</Label><Text>{props.complianceRisk}</Text></div>
          <div className={s.reviewItem}><Label>Read Timeframe</Label><Text>{props.readTimeframe}</Text></div>
          <div className={s.reviewItem}><Label>Acknowledgement Required</Label><Text>{props.requiresAcknowledgement ? 'Yes' : 'No'}</Text></div>
          <div className={s.reviewItem}><Label>Quiz Required</Label><Text>{props.requiresQuiz ? 'Yes' : 'No'}</Text></div>
        </div>
      )},
      { key: 'audience', icon: 'People', title: 'Target Audience', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Audience</Label><Text>{props.targetAllEmployees ? 'All Employees' : 'Specific groups'}</Text></div>
          {!props.targetAllEmployees && <>
            <div className={s.reviewItem}><Label>Departments</Label><Text>{props.targetDepartments.join(', ') || 'None'}</Text></div>
            <div className={s.reviewItem}><Label>Roles</Label><Text>{props.targetRoles.join(', ') || 'None'}</Text></div>
            <div className={s.reviewItem}><Label>Locations</Label><Text>{props.targetLocations.join(', ') || 'None'}</Text></div>
          </>}
        </div>
      )},
      { key: 'dates', icon: 'Calendar', title: 'Dates', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Effective Date</Label><Text>{props.effectiveDate || '-'}</Text></div>
          <div className={s.reviewItem}><Label>Expiry Date</Label><Text>{props.expiryDate || 'No expiry'}</Text></div>
          <div className={s.reviewItem}><Label>Review Frequency</Label><Text>{props.reviewFrequency}</Text></div>
        </div>
      )},
      { key: 'workflow', icon: 'Flow', title: 'Workflow', content: (
        <div className={s.reviewGrid}>
          <div className={s.reviewItem}><Label>Reviewers</Label><Text>{props.reviewers.length > 0 ? `${props.reviewers.length} assigned` : 'None'}</Text></div>
          <div className={s.reviewItem}><Label>Approvers</Label><Text>{props.approvers.length > 0 ? `${props.approvers.length} assigned` : 'None'}</Text></div>
        </div>
      )},
    ];

    return (
      <div className={s.wizardStepContent}>
        <div className={s.reviewSummary}>
          {sections.map(section => {
            const isExpanded = props.expandedReviewSections.has(section.key);
            return (
              <div key={section.key} className={s.reviewSectionCollapsible}>
                <div className={s.reviewSectionToggle} onClick={() => toggleSection(section.key)}>
                  <Text variant="mediumPlus" className={s.reviewSectionTitle}>
                    <Icon iconName={section.icon} style={{ marginRight: 8 }} />{section.title}
                  </Text>
                  <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 12, color: '#6b7280' }} />
                </div>
                <div className={isExpanded ? s.reviewSectionBody : s.reviewSectionBodyCollapsed}>{section.content}</div>
              </div>
            );
          })}
        </div>
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginTop: 16 } }}>
          Please review all information carefully before submitting.
        </MessageBar>
      </div>
    );
  };

  // ──── Current Step Dispatcher ────
  const renderCurrentStep = (): JSX.Element => {
    switch (props.currentStep) {
      case 0: return renderStep0();
      case 3: return renderStep3();
      case 4: return renderStep4();
      case 5: return renderStep5();
      case 7: return renderStep7();
      // Steps 1, 2, 6 delegate to content editors / reviewers rendered by parent
      default: return <></>;
    }
  };

  // ──── Main 3-Column Layout ────
  return (
    <div className={s.v3Layout}>
      {renderSidebar()}
      <main className={s.v3Main}>
        <div className={s.v3StepHeader}>
          <Text variant="xLarge" style={{ fontWeight: 700, color: '#0f172a' }}>
            {WIZARD_STEPS[props.currentStep]?.title || 'Step'}
          </Text>
          <Text variant="small" style={{ color: '#64748b', marginTop: 4 }}>
            Step {props.currentStep + 1} of {WIZARD_STEPS.length}
          </Text>
        </div>
        {renderCurrentStep()}
        {/* Navigation buttons */}
        <Stack horizontal horizontalAlign="space-between" style={{ marginTop: 24, paddingTop: 16, borderTop: '1px solid #e2e8f0' }}>
          <DefaultButton
            text="Previous"
            iconProps={{ iconName: 'ChevronLeft' }}
            onClick={props.onPreviousStep}
            disabled={props.currentStep === 0}
          />
          {props.currentStep < WIZARD_STEPS.length - 1 ? (
            <PrimaryButton
              text="Next"
              iconProps={{ iconName: 'ChevronRight' }}
              onClick={props.onNextStep}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
            />
          ) : (
            <PrimaryButton
              text="Submit for Review"
              iconProps={{ iconName: 'Send' }}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
            />
          )}
        </Stack>
      </main>
      {renderContextPanel()}
    </div>
  );
};

export default PolicyWizard;
