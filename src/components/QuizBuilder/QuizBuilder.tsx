import * as React from 'react';
import {
  Stack,
  Text,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Toggle,
  SpinnerSize,
  Spinner,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  CommandBar,
  ICommandBarItemProps,
  Dialog,
  DialogType,
  DialogFooter,
  Label,
  Slider,
  mergeStyleSets,
  Pivot,
  PivotItem,
  DatePicker,
  Checkbox,
  ChoiceGroup,
  ActionButton,
  Separator
} from '@fluentui/react';
import {
  QuizService,
  IQuiz,
  IQuizQuestion,
  QuestionType,
  DifficultyLevel,
  QuizStatus,
  IQuestionBank,
  IQuizSection,
  IQuizStatistics,
  IQuizExportData
} from '../../services/QuizService';
import { PolicyService } from '../../services/PolicyService';
import { IPolicy, PolicyStatus } from '../../models/IPolicy';
import { SPFI } from '@pnp/sp';

export interface IQuizBuilderProps {
  sp: SPFI;
  quizId?: number;
  policyId?: number;
  onSave?: (quiz: IQuiz) => void;
  onCancel?: () => void;
}

// Enhanced state interface
export interface IQuizBuilderState {
  loading: boolean;
  saving: boolean;
  error: string | null;
  success: string | null;
  activeTab: string;

  // Quiz details
  quizId: number | null;
  title: string;
  description: string;
  policyId: number | null;
  passingScore: number;
  timeLimit: number;
  maxAttempts: number;
  status: QuizStatus;
  category: string;
  difficultyLevel: DifficultyLevel;
  randomizeQuestions: boolean;
  showCorrectAnswers: boolean;

  // Scheduling
  scheduledStartDate: Date | null;
  scheduledEndDate: Date | null;
  isScheduled: boolean;

  // Advanced settings
  allowPartialCredit: boolean;
  shuffleWithinSections: boolean;
  requireSequentialCompletion: boolean;
  allowReview: boolean;

  // Certificate settings
  generateCertificate: boolean;
  certificateTemplateId: number | null;

  // Questions
  questions: IQuizQuestion[];

  // Sections
  sections: IQuizSection[];
  enableSections: boolean;

  // Policies for dropdown
  policies: IPolicy[];

  // Question banks
  questionBanks: IQuestionBank[];
  showQuestionBankPanel: boolean;
  selectedBankId: number | null;

  // Question editor
  showQuestionPanel: boolean;
  editingQuestion: Partial<IQuizQuestion> | null;
  editingQuestionIndex: number;

  // Delete confirmation
  showDeleteDialog: boolean;
  questionToDelete: number | null;

  // Import/Export
  showImportDialog: boolean;
  importFile: File | null;
  importFormat: 'json' | 'csv';
  showExportDialog: boolean;
  exportFormat: 'json' | 'csv';

  // Statistics
  showStatisticsPanel: boolean;
  quizStatistics: IQuizStatistics | null;

  // Section editor
  showSectionPanel: boolean;
  editingSection: Partial<IQuizSection> | null;
  editingSectionIndex: number;

  // Matching items editor
  matchingItems: Array<{ left: string; right: string }>;

  // Ordering items editor
  orderingItems: string[];

  // Fill in blank items
  blankAnswers: string[];

  // Hotspot regions
  hotspotRegions: Array<{ x: number; y: number; width: number; height: number; label: string }>;

  // Image choices
  imageChoices: Array<{ imageUrl: string; label: string; isCorrect: boolean }>;

  // Rating scale
  ratingMin: number;
  ratingMax: number;
  ratingLabels: { min: string; max: string };

  // Grading rubric for essays
  gradingRubric: Array<{ criteria: string; maxPoints: number; description: string }>;

  // AI Question Builder
  showAiPanel: boolean;
  aiQuestionCount: number;
  aiDifficulty: string;
  aiIncludeExcerpts: boolean;
  aiSelectedTypes: string[];
  aiGenerating: boolean;
  aiError: string | null;
  aiGeneratedQuestions: IQuizQuestion[];
  aiFunctionUrl: string;
}

const styles = mergeStyleSets({
  container: {
    padding: '20px',
    maxWidth: '1400px',
    margin: '0 auto',
    backgroundColor: '#faf9f8'
  },
  header: {
    marginBottom: '24px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center'
  },
  title: {
    fontSize: '28px',
    fontWeight: 600,
    color: '#323130'
  },
  section: {
    backgroundColor: '#ffffff',
    padding: '20px',
    borderRadius: '8px',
    boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
    marginBottom: '20px'
  },
  sectionTitle: {
    fontSize: '18px',
    fontWeight: 600,
    marginBottom: '16px',
    color: '#323130',
    display: 'flex',
    alignItems: 'center',
    gap: '8px'
  },
  formRow: {
    marginBottom: '16px'
  },
  questionCard: {
    backgroundColor: '#f8f8f8',
    padding: '16px',
    borderRadius: '6px',
    marginBottom: '12px',
    border: '1px solid #edebe9',
    transition: 'box-shadow 0.2s',
    ':hover': {
      boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
    }
  },
  questionHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginBottom: '8px'
  },
  questionNumber: {
    backgroundColor: '#0078d4',
    color: 'white',
    borderRadius: '50%',
    width: '28px',
    height: '28px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '14px',
    fontWeight: 600,
    marginRight: '12px',
    flexShrink: 0
  },
  questionText: {
    flex: 1,
    fontSize: '15px',
    fontWeight: 500,
    color: '#323130'
  },
  questionType: {
    fontSize: '12px',
    color: '#605e5c',
    backgroundColor: '#e1dfdd',
    padding: '2px 8px',
    borderRadius: '4px',
    marginLeft: '8px'
  },
  questionTypeBadge: {
    fontSize: '11px',
    padding: '2px 8px',
    borderRadius: '12px',
    fontWeight: 500
  },
  optionsList: {
    marginTop: '12px',
    paddingLeft: '40px'
  },
  option: {
    padding: '4px 0',
    fontSize: '14px',
    color: '#605e5c'
  },
  correctOption: {
    color: '#107c10',
    fontWeight: 500
  },
  points: {
    fontSize: '12px',
    color: '#0078d4',
    fontWeight: 500
  },
  emptyState: {
    textAlign: 'center' as const,
    padding: '40px',
    color: '#605e5c'
  },
  panelContent: {
    padding: '20px 0'
  },
  optionInput: {
    marginBottom: '8px'
  },
  correctAnswerSection: {
    marginTop: '16px',
    padding: '12px',
    backgroundColor: '#f3f2f1',
    borderRadius: '4px'
  },
  advancedSection: {
    marginTop: '20px',
    padding: '16px',
    backgroundColor: '#f8f8f8',
    borderRadius: '8px',
    border: '1px dashed #d2d0ce'
  },
  scheduleSection: {
    backgroundColor: '#fff4ce',
    padding: '16px',
    borderRadius: '8px',
    border: '1px solid #ffb900',
    marginTop: '16px'
  },
  certificateSection: {
    backgroundColor: '#dff6dd',
    padding: '16px',
    borderRadius: '8px',
    border: '1px solid #107c10',
    marginTop: '16px'
  },
  statisticsCard: {
    backgroundColor: '#f0f6ff',
    padding: '16px',
    borderRadius: '8px',
    marginBottom: '12px'
  },
  statValue: {
    fontSize: '32px',
    fontWeight: 700,
    color: '#0078d4'
  },
  statLabel: {
    fontSize: '12px',
    color: '#605e5c',
    textTransform: 'uppercase' as const
  },
  matchingPairRow: {
    display: 'flex',
    gap: '16px',
    alignItems: 'center',
    marginBottom: '8px',
    padding: '8px',
    backgroundColor: '#f3f2f1',
    borderRadius: '4px'
  },
  orderingItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    backgroundColor: '#fff',
    border: '1px solid #d2d0ce',
    borderRadius: '4px',
    marginBottom: '4px',
    cursor: 'grab'
  },
  hotspotContainer: {
    position: 'relative' as const,
    display: 'inline-block',
    border: '2px dashed #0078d4',
    borderRadius: '4px'
  },
  hotspotRegion: {
    position: 'absolute' as const,
    border: '2px solid #107c10',
    backgroundColor: 'rgba(16, 124, 16, 0.2)',
    borderRadius: '4px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#107c10',
    fontWeight: 600,
    fontSize: '12px'
  },
  imageChoiceCard: {
    width: '150px',
    padding: '8px',
    border: '2px solid #d2d0ce',
    borderRadius: '8px',
    textAlign: 'center' as const,
    cursor: 'pointer',
    transition: 'border-color 0.2s'
  },
  imageChoiceSelected: {
    borderColor: '#107c10',
    backgroundColor: '#dff6dd'
  },
  rubricRow: {
    padding: '12px',
    backgroundColor: '#f8f8f8',
    borderRadius: '4px',
    marginBottom: '8px'
  },
  importExportButtons: {
    display: 'flex',
    gap: '8px'
  },
  badge: {
    display: 'inline-block',
    padding: '2px 8px',
    borderRadius: '12px',
    fontSize: '11px',
    fontWeight: 500
  },
  badgeSuccess: {
    backgroundColor: '#dff6dd',
    color: '#107c10'
  },
  badgeWarning: {
    backgroundColor: '#fff4ce',
    color: '#d83b01'
  },
  badgeInfo: {
    backgroundColor: '#cce4f6',
    color: '#0078d4'
  },
  tabContent: {
    padding: '20px 0'
  },
  sectionCard: {
    backgroundColor: '#f0f6ff',
    padding: '16px',
    borderRadius: '8px',
    marginBottom: '12px',
    border: '1px solid #0078d4'
  },
  dragHandle: {
    cursor: 'grab',
    color: '#605e5c',
    marginRight: '8px'
  }
});

// Question type options - all 11 types
const questionTypeOptions: IDropdownOption[] = [
  { key: QuestionType.MultipleChoice, text: 'Multiple Choice', data: { icon: 'RadioBtnOn' } },
  { key: QuestionType.TrueFalse, text: 'True/False', data: { icon: 'CheckMark' } },
  { key: QuestionType.MultipleSelect, text: 'Multiple Select', data: { icon: 'CheckboxComposite' } },
  { key: QuestionType.ShortAnswer, text: 'Short Answer', data: { icon: 'TextField' } },
  { key: QuestionType.FillInBlank, text: 'Fill in the Blank', data: { icon: 'TextBox' } },
  { key: QuestionType.Matching, text: 'Matching', data: { icon: 'Link' } },
  { key: QuestionType.Ordering, text: 'Ordering/Sequencing', data: { icon: 'Sort' } },
  { key: QuestionType.RatingScale, text: 'Rating Scale', data: { icon: 'FavoriteStar' } },
  { key: QuestionType.Essay, text: 'Essay/Long Answer', data: { icon: 'EditNote' } },
  { key: QuestionType.ImageChoice, text: 'Image-based Choice', data: { icon: 'Photo2' } },
  { key: QuestionType.Hotspot, text: 'Hotspot/Image Map', data: { icon: 'MapPin' } }
];

const difficultyOptions: IDropdownOption[] = [
  { key: DifficultyLevel.Easy, text: 'Easy' },
  { key: DifficultyLevel.Medium, text: 'Medium' },
  { key: DifficultyLevel.Hard, text: 'Hard' },
  { key: DifficultyLevel.Expert, text: 'Expert' }
];

const statusOptions: IDropdownOption[] = [
  { key: QuizStatus.Draft, text: 'Draft' },
  { key: QuizStatus.Published, text: 'Published' },
  { key: QuizStatus.Scheduled, text: 'Scheduled' },
  { key: QuizStatus.Archived, text: 'Archived' }
];

const categoryOptions: IDropdownOption[] = [
  { key: 'HR Policies', text: 'HR Policies' },
  { key: 'IT & Security', text: 'IT & Security' },
  { key: 'Health & Safety', text: 'Health & Safety' },
  { key: 'Compliance', text: 'Compliance' },
  { key: 'Data Privacy', text: 'Data Privacy' },
  { key: 'Financial', text: 'Financial' },
  { key: 'Onboarding', text: 'Onboarding' },
  { key: 'Training', text: 'Training' },
  { key: 'General', text: 'General' }
];

export class QuizBuilder extends React.Component<IQuizBuilderProps, IQuizBuilderState> {
  private quizService: QuizService;
  private policyService: PolicyService;
  private fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IQuizBuilderProps) {
    super(props);

    this.quizService = new QuizService(props.sp);
    this.policyService = new PolicyService(props.sp);
    this.fileInputRef = React.createRef();

    this.state = {
      loading: true,
      saving: false,
      error: null,
      success: null,
      activeTab: 'settings',

      quizId: props.quizId || null,
      title: '',
      description: '',
      policyId: props.policyId || null,
      passingScore: 70,
      timeLimit: 15,
      maxAttempts: 3,
      status: QuizStatus.Draft,
      category: 'General',
      difficultyLevel: DifficultyLevel.Medium,
      randomizeQuestions: true,
      showCorrectAnswers: true,

      // Scheduling
      scheduledStartDate: null,
      scheduledEndDate: null,
      isScheduled: false,

      // Advanced
      allowPartialCredit: true,
      shuffleWithinSections: false,
      requireSequentialCompletion: false,
      allowReview: true,

      // Certificate
      generateCertificate: false,
      certificateTemplateId: null,

      questions: [],
      sections: [],
      enableSections: false,
      policies: [],
      questionBanks: [],
      showQuestionBankPanel: false,
      selectedBankId: null,
      showQuestionPanel: false,
      editingQuestion: null,
      editingQuestionIndex: -1,
      showDeleteDialog: false,
      questionToDelete: null,
      showImportDialog: false,
      importFile: null,
      importFormat: 'json',
      showExportDialog: false,
      exportFormat: 'json',
      showStatisticsPanel: false,
      quizStatistics: null,
      showSectionPanel: false,
      editingSection: null,
      editingSectionIndex: -1,

      // Question type editors
      matchingItems: [{ left: '', right: '' }],
      orderingItems: [''],
      blankAnswers: [''],
      hotspotRegions: [],
      imageChoices: [],
      ratingMin: 1,
      ratingMax: 5,
      ratingLabels: { min: 'Poor', max: 'Excellent' },
      gradingRubric: [],

      // AI Question Builder
      showAiPanel: false,
      aiQuestionCount: 10,
      aiDifficulty: 'Medium',
      aiIncludeExcerpts: true,
      aiSelectedTypes: ['Multiple Choice', 'True/False', 'Multiple Select', 'Short Answer'],
      aiGenerating: false,
      aiError: null,
      aiGeneratedQuestions: [],
      aiFunctionUrl: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });

      await this.policyService.initialize();

      // Load policies for dropdown
      const policies = await this.policyService.getPolicies({ status: PolicyStatus.Published });

      // Load question banks
      const questionBanks = await this.quizService.getQuestionBanks();

      // If editing existing quiz, load it
      if (this.props.quizId) {
        const quiz = await this.quizService.getQuizById(this.props.quizId);
        const questions = await this.quizService.getQuizQuestions(this.props.quizId);

        if (quiz) {
          this.setState({
            quizId: quiz.Id,
            title: quiz.Title,
            description: quiz.QuizDescription || '',
            policyId: quiz.PolicyId,
            passingScore: quiz.PassingScore,
            timeLimit: quiz.TimeLimit,
            maxAttempts: quiz.MaxAttempts,
            status: quiz.Status as QuizStatus || QuizStatus.Draft,
            category: quiz.QuizCategory || 'General',
            difficultyLevel: quiz.DifficultyLevel as DifficultyLevel || DifficultyLevel.Medium,
            randomizeQuestions: quiz.RandomizeQuestions,
            showCorrectAnswers: quiz.ShowCorrectAnswers,
            isScheduled: quiz.Status === QuizStatus.Scheduled,
            scheduledStartDate: quiz.ScheduledStartDate ? new Date(quiz.ScheduledStartDate) : null,
            scheduledEndDate: quiz.ScheduledEndDate ? new Date(quiz.ScheduledEndDate) : null,
            allowPartialCredit: quiz.AllowPartialCredit ?? true,
            shuffleWithinSections: quiz.ShuffleWithinSections ?? false,
            requireSequentialCompletion: quiz.RequireSequentialCompletion ?? false,
            allowReview: quiz.AllowReview ?? true,
            generateCertificate: quiz.GenerateCertificate ?? false,
            certificateTemplateId: quiz.CertificateTemplateId ?? null,
            questions
          });
        }
      }

      this.setState({ policies, questionBanks, loading: false });
    } catch (error) {
      console.error('Failed to load data:', error);
      this.setState({
        error: 'Failed to load data. Please try again.',
        loading: false
      });
    }
  }

  private handleSaveQuiz = async (): Promise<void> => {
    const {
      quizId,
      title,
      description,
      policyId,
      passingScore,
      timeLimit,
      maxAttempts,
      status,
      category,
      difficultyLevel,
      randomizeQuestions,
      showCorrectAnswers,
      isScheduled,
      scheduledStartDate,
      scheduledEndDate,
      allowPartialCredit,
      shuffleWithinSections,
      requireSequentialCompletion,
      allowReview,
      generateCertificate,
      certificateTemplateId,
      policies
    } = this.state;

    if (!title.trim()) {
      this.setState({ error: 'Please enter a quiz title.' });
      return;
    }

    try {
      this.setState({ saving: true, error: null, success: null });

      const selectedPolicy = policies.find(p => p.Id === policyId);

      const quizData: Partial<IQuiz> = {
        Title: title,
        QuizDescription: description,
        PolicyId: policyId || undefined,
        PolicyTitle: selectedPolicy?.PolicyName || '',
        PassingScore: passingScore,
        TimeLimit: timeLimit,
        MaxAttempts: maxAttempts,
        Status: isScheduled ? QuizStatus.Scheduled : status,
        IsActive: status === QuizStatus.Published || status === QuizStatus.Scheduled,
        QuizCategory: category,
        DifficultyLevel: difficultyLevel,
        RandomizeQuestions: randomizeQuestions,
        ShowCorrectAnswers: showCorrectAnswers,
        ScheduledStartDate: isScheduled && scheduledStartDate ? scheduledStartDate.toISOString() : undefined,
        ScheduledEndDate: isScheduled && scheduledEndDate ? scheduledEndDate.toISOString() : undefined,
        AllowPartialCredit: allowPartialCredit,
        ShuffleWithinSections: shuffleWithinSections,
        RequireSequentialCompletion: requireSequentialCompletion,
        AllowReview: allowReview,
        GenerateCertificate: generateCertificate,
        CertificateTemplateId: certificateTemplateId || undefined
      };

      let savedQuiz: IQuiz | null;

      if (quizId) {
        await this.quizService.updateQuiz(quizId, quizData);
        savedQuiz = await this.quizService.getQuizById(quizId);
      } else {
        savedQuiz = await this.quizService.createQuiz(quizData);
        if (savedQuiz) {
          this.setState({ quizId: savedQuiz.Id });
        }
      }

      this.setState({
        saving: false,
        success: quizId ? 'Quiz updated successfully!' : 'Quiz created successfully!'
      });

      if (savedQuiz && this.props.onSave) {
        this.props.onSave(savedQuiz);
      }
    } catch (error) {
      console.error('Failed to save quiz:', error);
      this.setState({
        saving: false,
        error: 'Failed to save quiz. Please try again.'
      });
    }
  };

  private handleAddQuestion = (): void => {
    this.setState({
      showQuestionPanel: true,
      editingQuestion: {
        QuestionType: QuestionType.MultipleChoice,
        Points: 10,
        DifficultyLevel: DifficultyLevel.Medium,
        IsActive: true,
        IsRequired: true,
        PartialCreditEnabled: false,
        NegativeMarking: false,
        OptionA: '',
        OptionB: '',
        OptionC: '',
        OptionD: '',
        CorrectAnswer: 'A',
        Explanation: '',
        TimeLimit: 0,
        Tags: '',
        Hint: '',
        QuestionImage: ''
      },
      editingQuestionIndex: -1,
      matchingItems: [{ left: '', right: '' }],
      orderingItems: [''],
      blankAnswers: [''],
      hotspotRegions: [],
      imageChoices: [],
      ratingMin: 1,
      ratingMax: 5,
      ratingLabels: { min: 'Poor', max: 'Excellent' },
      gradingRubric: []
    });
  };

  private handleEditQuestion = (question: IQuizQuestion, index: number): void => {
    // Parse JSON fields for specific question types
    let matchingItems = [{ left: '', right: '' }];
    let orderingItems = [''];
    let blankAnswers = [''];
    let hotspotRegions: Array<{ x: number; y: number; width: number; height: number; label: string }> = [];
    let imageChoices: Array<{ imageUrl: string; label: string; isCorrect: boolean }> = [];
    let gradingRubric: Array<{ criteria: string; maxPoints: number; description: string }> = [];
    let ratingMin = 1;
    let ratingMax = 5;
    let ratingLabels = { min: 'Poor', max: 'Excellent' };

    try {
      if (question.QuestionType === QuestionType.Matching && question.MatchingPairs) {
        matchingItems = JSON.parse(question.MatchingPairs);
      }
      if (question.QuestionType === QuestionType.Ordering && question.OrderingItems) {
        orderingItems = JSON.parse(question.OrderingItems);
      }
      if (question.QuestionType === QuestionType.FillInBlank && question.BlankAnswers) {
        blankAnswers = JSON.parse(question.BlankAnswers);
      }
      if (question.QuestionType === QuestionType.Hotspot && question.HotspotData) {
        const hotspotData = JSON.parse(question.HotspotData);
        hotspotRegions = hotspotData.regions || [];
      }
      if (question.QuestionType === QuestionType.ImageChoice) {
        // Build from individual option images
        imageChoices = [
          question.OptionAImage ? { imageUrl: question.OptionAImage, label: question.OptionA || 'A', isCorrect: question.CorrectAnswer?.includes('A') || false } : null,
          question.OptionBImage ? { imageUrl: question.OptionBImage, label: question.OptionB || 'B', isCorrect: question.CorrectAnswer?.includes('B') || false } : null,
          question.OptionCImage ? { imageUrl: question.OptionCImage, label: question.OptionC || 'C', isCorrect: question.CorrectAnswer?.includes('C') || false } : null,
          question.OptionDImage ? { imageUrl: question.OptionDImage, label: question.OptionD || 'D', isCorrect: question.CorrectAnswer?.includes('D') || false } : null,
        ].filter((c): c is { imageUrl: string; label: string; isCorrect: boolean } => c !== null);
      }
      if (question.QuestionType === QuestionType.RatingScale) {
        ratingMin = question.ScaleMin || 1;
        ratingMax = question.ScaleMax || 5;
        if (question.ScaleLabels) {
          ratingLabels = JSON.parse(question.ScaleLabels);
        }
      }
    } catch (e) {
      console.error('Error parsing question data:', e);
    }

    this.setState({
      showQuestionPanel: true,
      editingQuestion: { ...question },
      editingQuestionIndex: index,
      matchingItems,
      orderingItems,
      blankAnswers,
      hotspotRegions,
      imageChoices,
      gradingRubric,
      ratingMin,
      ratingMax,
      ratingLabels
    });
  };

  private handleSaveQuestion = async (): Promise<void> => {
    const {
      quizId,
      editingQuestion,
      editingQuestionIndex,
      questions,
      matchingItems,
      orderingItems,
      blankAnswers,
      hotspotRegions,
      imageChoices,
      // gradingRubric used in essay type editor panel — not needed in saveQuestion
      ratingMin,
      ratingMax,
      ratingLabels
    } = this.state;

    if (!editingQuestion?.QuestionText?.trim()) {
      this.setState({ error: 'Please enter the question text.' });
      return;
    }

    if (!quizId) {
      this.setState({ error: 'Please save the quiz first before adding questions.' });
      return;
    }

    try {
      this.setState({ saving: true, error: null });

      // Prepare question data based on type
      const questionData: Partial<IQuizQuestion> = {
        ...editingQuestion,
        QuizId: quizId,
        QuestionOrder: editingQuestionIndex >= 0 ? questions[editingQuestionIndex].QuestionOrder : questions.length + 1
      };

      // Add type-specific data
      switch (editingQuestion.QuestionType) {
        case QuestionType.Matching:
          questionData.MatchingPairs = JSON.stringify(matchingItems.filter(m => m.left && m.right));
          break;
        case QuestionType.Ordering:
          questionData.OrderingItems = JSON.stringify(orderingItems.filter(o => o.trim()));
          questionData.CorrectAnswer = JSON.stringify(orderingItems.filter(o => o.trim()));
          break;
        case QuestionType.FillInBlank:
          questionData.BlankAnswers = JSON.stringify(blankAnswers.filter(b => b.trim()));
          break;
        case QuestionType.Hotspot:
          questionData.HotspotData = JSON.stringify({
            imageUrl: editingQuestion.QuestionImage || '',
            regions: hotspotRegions
          });
          break;
        case QuestionType.ImageChoice:
          // Map to individual option images
          imageChoices.forEach((choice, idx) => {
            const letter = String.fromCharCode(65 + idx); // A, B, C, D
            (questionData as Record<string, unknown>)[`Option${letter}Image`] = choice.imageUrl;
            (questionData as Record<string, unknown>)[`Option${letter}`] = choice.label;
          });
          // Set correct answers
          questionData.CorrectAnswer = imageChoices
            .map((c, i) => c.isCorrect ? String.fromCharCode(65 + i) : null)
            .filter(Boolean)
            .join(';');
          break;
        case QuestionType.Essay:
          // Essay questions are graded manually by rubric reference
          break;
        case QuestionType.RatingScale:
          questionData.ScaleMin = ratingMin;
          questionData.ScaleMax = ratingMax;
          questionData.ScaleLabels = JSON.stringify(ratingLabels);
          break;
      }

      if (editingQuestionIndex >= 0 && editingQuestion.Id) {
        await this.quizService.updateQuestion(editingQuestion.Id, questionData);
      } else {
        await this.quizService.createQuestion(questionData);
      }

      const updatedQuestions = await this.quizService.getQuizQuestions(quizId);

      this.setState({
        questions: updatedQuestions,
        showQuestionPanel: false,
        editingQuestion: null,
        editingQuestionIndex: -1,
        saving: false,
        success: 'Question saved successfully!'
      });
    } catch (error) {
      console.error('Failed to save question:', error);
      this.setState({
        saving: false,
        error: 'Failed to save question. Please try again.'
      });
    }
  };

  private handleDeleteQuestion = async (): Promise<void> => {
    const { quizId, questionToDelete } = this.state;

    if (!questionToDelete || !quizId) return;

    try {
      this.setState({ saving: true, error: null });

      await this.quizService.deleteQuestion(questionToDelete, quizId);

      const updatedQuestions = await this.quizService.getQuizQuestions(quizId);

      this.setState({
        questions: updatedQuestions,
        showDeleteDialog: false,
        questionToDelete: null,
        saving: false,
        success: 'Question deleted successfully!'
      });
    } catch (error) {
      console.error('Failed to delete question:', error);
      this.setState({
        saving: false,
        error: 'Failed to delete question. Please try again.'
      });
    }
  };

  private handleMoveQuestion = async (fromIndex: number, toIndex: number): Promise<void> => {
    const { questions, quizId } = this.state;
    if (!quizId || toIndex < 0 || toIndex >= questions.length) return;

    const updated = [...questions];
    const [moved] = updated.splice(fromIndex, 1);
    updated.splice(toIndex, 0, moved);

    // Update order in state immediately for responsiveness
    this.setState({ questions: updated });

    // Persist order to SharePoint
    try {
      for (let i = 0; i < updated.length; i++) {
        if (updated[i].QuestionOrder !== i + 1) {
          await this.quizService.updateQuestion(updated[i].Id, { QuestionOrder: i + 1 });
          updated[i] = { ...updated[i], QuestionOrder: i + 1 };
        }
      }
      this.setState({ questions: updated, success: 'Question order updated.' });
    } catch (err) {
      console.error('Failed to update question order:', err);
      this.setState({ error: 'Failed to save question order.' });
    }
  };

  private handleDuplicateQuestion = async (question: IQuizQuestion): Promise<void> => {
    const { quizId, questions } = this.state;
    if (!quizId) return;

    try {
      this.setState({ saving: true });

      // Create a copy of the question without the Id
      const copy: Partial<IQuizQuestion> = {
        ...question,
        QuestionText: `${question.QuestionText} (Copy)`,
        QuestionOrder: questions.length + 1,
        TimesAnswered: 0,
        TimesCorrect: 0,
        AverageTime: 0
      };
      delete (copy as Record<string, unknown>).Id;

      await this.quizService.createQuestion(copy);
      const updatedQuestions = await this.quizService.getQuizQuestions(quizId);

      this.setState({
        questions: updatedQuestions,
        saving: false,
        success: 'Question duplicated!'
      });
    } catch (err) {
      console.error('Failed to duplicate question:', err);
      this.setState({ saving: false, error: 'Failed to duplicate question.' });
    }
  };

  private handleImportQuestions = async (): Promise<void> => {
    const { quizId, importFile, importFormat } = this.state;

    if (!quizId || !importFile) {
      this.setState({ error: 'Please save the quiz first and select a file to import.' });
      return;
    }

    try {
      this.setState({ saving: true, error: null });

      const fileContent = await importFile.text();

      if (importFormat === 'json') {
        const data: IQuizExportData = JSON.parse(fileContent);
        for (const question of data.questions) {
          await this.quizService.createQuestion({
            ...question,
            QuizId: quizId,
            Id: undefined
          });
        }
      } else if (importFormat === 'csv') {
        await this.quizService.importQuestionsFromCSV(quizId, fileContent);
      }

      const updatedQuestions = await this.quizService.getQuizQuestions(quizId);

      this.setState({
        questions: updatedQuestions,
        showImportDialog: false,
        importFile: null,
        saving: false,
        success: `Questions imported successfully!`
      });
    } catch (error) {
      console.error('Failed to import questions:', error);
      this.setState({
        saving: false,
        error: 'Failed to import questions. Please check the file format.'
      });
    }
  };

  private handleExportQuestions = async (): Promise<void> => {
    const { quizId, questions, exportFormat, title } = this.state;

    if (!quizId) return;

    try {
      let content: string;
      let filename: string;
      let mimeType: string;

      if (exportFormat === 'json') {
        const exportData: IQuizExportData = {
          version: '1.0',
          exportDate: new Date().toISOString(),
          quiz: { Title: title },
          sections: [],
          questions: questions
        };
        content = JSON.stringify(exportData, null, 2);
        filename = `${title.replace(/\s+/g, '_')}_questions.json`;
        mimeType = 'application/json';
      } else {
        content = await this.quizService.exportQuestionsToCSV(questions);
        filename = `${title.replace(/\s+/g, '_')}_questions.csv`;
        mimeType = 'text/csv';
      }

      // Create and download file
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      this.setState({
        showExportDialog: false,
        success: 'Questions exported successfully!'
      });
    } catch (error) {
      console.error('Failed to export questions:', error);
      this.setState({
        error: 'Failed to export questions. Please try again.'
      });
    }
  };

  private handleLoadStatistics = async (): Promise<void> => {
    const { quizId } = this.state;

    if (!quizId) return;

    try {
      const stats = await this.quizService.getQuizStatistics(quizId);
      this.setState({
        quizStatistics: stats,
        showStatisticsPanel: true
      });
    } catch (error) {
      console.error('Failed to load statistics:', error);
      this.setState({ error: 'Failed to load quiz statistics.' });
    }
  };

  private renderCommandBar(): JSX.Element {
    const { saving, quizId } = this.state;

    const items: ICommandBarItemProps[] = [
      {
        key: 'save',
        text: quizId ? 'Update Quiz' : 'Create Quiz',
        iconProps: { iconName: 'Save' },
        onClick: (): void => { void this.handleSaveQuiz(); },
        disabled: saving
      },
      {
        key: 'addQuestion',
        text: 'Add Question',
        iconProps: { iconName: 'Add' },
        onClick: this.handleAddQuestion,
        disabled: saving || !quizId
      },
      {
        key: 'aiGenerate',
        text: 'AI Generate',
        iconProps: { iconName: 'Robot' },
        onClick: () => this.setState({ showAiPanel: true }),
        disabled: saving || !quizId
      }
    ];

    const farItems: ICommandBarItemProps[] = [
      {
        key: 'import',
        text: 'Import',
        iconProps: { iconName: 'Upload' },
        onClick: () => this.setState({ showImportDialog: true }),
        disabled: !quizId
      },
      {
        key: 'export',
        text: 'Export',
        iconProps: { iconName: 'Download' },
        onClick: () => this.setState({ showExportDialog: true }),
        disabled: !quizId || this.state.questions.length === 0
      },
      {
        key: 'statistics',
        text: 'Statistics',
        iconProps: { iconName: 'BarChartVertical' },
        onClick: (): void => { void this.handleLoadStatistics(); },
        disabled: !quizId
      }
    ];

    if (this.props.onCancel) {
      farItems.push({
        key: 'cancel',
        text: 'Cancel',
        iconProps: { iconName: 'Cancel' },
        onClick: this.props.onCancel
      });
    }

    return <CommandBar items={items} farItems={farItems} />;
  }

  private renderQuizSettings(): JSX.Element {
    const {
      title,
      description,
      policyId,
      passingScore,
      timeLimit,
      maxAttempts,
      status,
      category,
      difficultyLevel,
      randomizeQuestions,
      showCorrectAnswers,
      isScheduled,
      scheduledStartDate,
      scheduledEndDate,
      policies
    } = this.state;

    const policyOptions: IDropdownOption[] = [
      { key: '', text: '(No linked policy)' },
      ...policies.map(p => ({ key: p.Id || 0, text: `${p.PolicyNumber} - ${p.PolicyName}` }))
    ];

    return (
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>
          <IconButton iconProps={{ iconName: 'Settings' }} disabled />
          Quiz Settings
        </Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal tokens={{ childrenGap: 16 }}>
            <Stack.Item grow={2}>
              <TextField
                label="Quiz Title"
                value={title}
                onChange={(e, value) => this.setState({ title: value || '' })}
                required
                placeholder="Enter quiz title..."
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <Dropdown
                label="Status"
                options={statusOptions}
                selectedKey={status}
                onChange={(e, option) => this.setState({ status: option?.key as QuizStatus || QuizStatus.Draft })}
              />
            </Stack.Item>
          </Stack>

          <TextField
            label="Description"
            value={description}
            onChange={(e, value) => this.setState({ description: value || '' })}
            multiline
            rows={3}
            placeholder="Enter quiz description..."
          />

          <Stack horizontal tokens={{ childrenGap: 16 }}>
            <Stack.Item grow={1}>
              <Dropdown
                label="Category"
                options={categoryOptions}
                selectedKey={category}
                onChange={(e, option) => this.setState({ category: option?.key as string || 'General' })}
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <Dropdown
                label="Linked Policy"
                options={policyOptions}
                selectedKey={policyId || ''}
                onChange={(e, option) => this.setState({ policyId: option?.key ? Number(option.key) : null })}
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <Dropdown
                label="Difficulty Level"
                options={difficultyOptions}
                selectedKey={difficultyLevel}
                onChange={(e, option) => this.setState({ difficultyLevel: option?.key as DifficultyLevel || DifficultyLevel.Medium })}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 16 }}>
            <Stack.Item grow={1}>
              <Label>Passing Score: {passingScore}%</Label>
              <Slider
                min={50}
                max={100}
                step={5}
                value={passingScore}
                onChange={(value) => this.setState({ passingScore: value })}
                showValue={false}
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <TextField
                label="Time Limit (minutes)"
                type="number"
                value={String(timeLimit)}
                onChange={(e, value) => this.setState({ timeLimit: parseInt(value || '15', 10) })}
                min={1}
                max={180}
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <TextField
                label="Max Attempts"
                type="number"
                value={String(maxAttempts)}
                onChange={(e, value) => this.setState({ maxAttempts: parseInt(value || '3', 10) })}
                min={1}
                max={10}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 32 }}>
            <Toggle
              label="Randomize Questions"
              checked={randomizeQuestions}
              onChange={(e, checked) => this.setState({ randomizeQuestions: checked || false })}
            />
            <Toggle
              label="Show Correct Answers"
              checked={showCorrectAnswers}
              onChange={(e, checked) => this.setState({ showCorrectAnswers: checked || false })}
            />
          </Stack>

          {/* Scheduling Section */}
          <Toggle
            label="Schedule Quiz"
            checked={isScheduled}
            onChange={(e, checked) => this.setState({ isScheduled: checked || false })}
          />

          {isScheduled && (
            <div className={styles.scheduleSection}>
              <Text className={styles.sectionTitle}>
                <IconButton iconProps={{ iconName: 'Calendar' }} disabled />
                Quiz Schedule
              </Text>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Stack.Item grow={1}>
                  <DatePicker
                    label="Start Date"
                    value={scheduledStartDate || undefined}
                    onSelectDate={(date) => this.setState({ scheduledStartDate: date || null })}
                    placeholder="Select start date..."
                  />
                </Stack.Item>
                <Stack.Item grow={1}>
                  <DatePicker
                    label="End Date"
                    value={scheduledEndDate || undefined}
                    onSelectDate={(date) => this.setState({ scheduledEndDate: date || null })}
                    placeholder="Select end date..."
                    minDate={scheduledStartDate || undefined}
                  />
                </Stack.Item>
              </Stack>
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderAdvancedSettings(): JSX.Element {
    const {
      allowPartialCredit,
      shuffleWithinSections,
      requireSequentialCompletion,
      allowReview
    } = this.state;

    return (
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>
          <IconButton iconProps={{ iconName: 'Settings' }} disabled />
          Advanced Settings
        </Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
            <Toggle
              label="Allow Partial Credit"
              checked={allowPartialCredit}
              onChange={(e, checked) => this.setState({ allowPartialCredit: checked || false })}
              onText="Yes"
              offText="No"
            />
            <Toggle
              label="Allow Review Before Submit"
              checked={allowReview}
              onChange={(e, checked) => this.setState({ allowReview: checked || false })}
              onText="Yes"
              offText="No"
            />
          </Stack>

          <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
            <Toggle
              label="Shuffle Within Sections"
              checked={shuffleWithinSections}
              onChange={(e, checked) => this.setState({ shuffleWithinSections: checked || false })}
              onText="Yes"
              offText="No"
            />
            <Toggle
              label="Require Sequential Completion"
              checked={requireSequentialCompletion}
              onChange={(e, checked) => this.setState({ requireSequentialCompletion: checked || false })}
              onText="Yes"
              offText="No"
            />
          </Stack>

          <MessageBar messageBarType={MessageBarType.info}>
            Additional quiz settings can be configured per-question, including negative marking and time limits.
          </MessageBar>
        </Stack>
      </div>
    );
  }

  private renderCertificateSettings(): JSX.Element {
    const {
      generateCertificate,
      title
    } = this.state;

    return (
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>
          <IconButton iconProps={{ iconName: 'Certificate' }} disabled />
          Certificate Settings
        </Text>

        <Toggle
          label="Generate Certificate on Pass"
          checked={generateCertificate}
          onChange={(e, checked) => this.setState({ generateCertificate: checked || false })}
          onText="Yes"
          offText="No"
        />

        {generateCertificate && (
          <div className={styles.certificateSection}>
            <Stack tokens={{ childrenGap: 16 }}>
              <MessageBar messageBarType={MessageBarType.success}>
                A certificate will be automatically generated when users pass this quiz.
              </MessageBar>
              <Text variant="medium">
                Certificate Title: <strong>Certificate of Completion - {title || 'Quiz Name'}</strong>
              </Text>
              <Text variant="small" style={{ color: '#605e5c' }}>
                Certificate templates can be managed in the admin settings. The default template will be used for this quiz.
              </Text>
            </Stack>
          </div>
        )}
      </div>
    );
  }

  private renderQuestions(): JSX.Element {
    const { questions, quizId } = this.state;

    return (
      <div className={styles.section}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className={styles.sectionTitle}>
            <IconButton iconProps={{ iconName: 'ClipboardList' }} disabled />
            Questions ({questions.length})
          </Text>
          {quizId && (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="From Question Bank"
                iconProps={{ iconName: 'Library' }}
                onClick={() => this.setState({ showQuestionBankPanel: true })}
              />
              <PrimaryButton
                text="Add Question"
                iconProps={{ iconName: 'Add' }}
                onClick={this.handleAddQuestion}
              />
            </Stack>
          )}
        </Stack>

        {!quizId && (
          <MessageBar messageBarType={MessageBarType.info}>
            Save the quiz first to add questions.
          </MessageBar>
        )}

        {quizId && questions.length === 0 && (
          <div className={styles.emptyState}>
            <Text variant="large">No questions yet</Text>
            <Text>Click "Add Question" to create your first question or import from a question bank.</Text>
          </div>
        )}

        {questions.map((question, index) => (
          <div key={question.Id} className={styles.questionCard}>
            <div className={styles.questionHeader}>
              <Stack horizontal verticalAlign="center" styles={{ root: { flex: 1 } }}>
                <div className={styles.questionNumber}>{index + 1}</div>
                <Text className={styles.questionText}>{question.QuestionText}</Text>
                <span
                  className={styles.questionTypeBadge}
                  style={{
                    backgroundColor: this.getQuestionTypeColor(question.QuestionType).bg,
                    color: this.getQuestionTypeColor(question.QuestionType).text
                  }}
                >
                  {question.QuestionType}
                </span>
                {question.QuestionType === QuestionType.Essay && (
                  <span className={`${styles.badge} ${styles.badgeWarning}`} style={{ marginLeft: '8px' }}>
                    Manual Grading
                  </span>
                )}
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                <Text className={styles.points}>{question.Points} pts</Text>
                <IconButton
                  iconProps={{ iconName: 'Up' }}
                  title="Move Up"
                  disabled={index === 0}
                  onClick={() => this.handleMoveQuestion(index, index - 1)}
                  styles={{ root: { height: 28, width: 28 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Down' }}
                  title="Move Down"
                  disabled={index === questions.length - 1}
                  onClick={() => this.handleMoveQuestion(index, index + 1)}
                  styles={{ root: { height: 28, width: 28 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Copy' }}
                  title="Duplicate"
                  onClick={() => this.handleDuplicateQuestion(question)}
                />
                <IconButton
                  iconProps={{ iconName: 'Edit' }}
                  title="Edit"
                  onClick={() => this.handleEditQuestion(question, index)}
                />
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Delete"
                  onClick={() => this.setState({ showDeleteDialog: true, questionToDelete: question.Id })}
                />
              </Stack>
            </div>

            {this.renderQuestionPreview(question)}
          </div>
        ))}
      </div>
    );
  }

  private getQuestionTypeColor(type: string): { bg: string; text: string } {
    const colors: Record<string, { bg: string; text: string }> = {
      [QuestionType.MultipleChoice]: { bg: '#cce4f6', text: '#0078d4' },
      [QuestionType.TrueFalse]: { bg: '#dff6dd', text: '#107c10' },
      [QuestionType.MultipleSelect]: { bg: '#e8daef', text: '#7b4397' },
      [QuestionType.ShortAnswer]: { bg: '#fef3cd', text: '#856404' },
      [QuestionType.FillInBlank]: { bg: '#d4edda', text: '#155724' },
      [QuestionType.Matching]: { bg: '#d1ecf1', text: '#0c5460' },
      [QuestionType.Ordering]: { bg: '#f8d7da', text: '#721c24' },
      [QuestionType.RatingScale]: { bg: '#fff3cd', text: '#856404' },
      [QuestionType.Essay]: { bg: '#e2e3e5', text: '#383d41' },
      [QuestionType.ImageChoice]: { bg: '#cce5ff', text: '#004085' },
      [QuestionType.Hotspot]: { bg: '#f5c6cb', text: '#721c24' }
    };
    return colors[type] || { bg: '#e1dfdd', text: '#605e5c' };
  }

  private renderQuestionPreview(question: IQuizQuestion): JSX.Element {
    switch (question.QuestionType) {
      case QuestionType.TrueFalse:
        return (
          <div className={styles.optionsList}>
            <div className={`${styles.option} ${question.CorrectAnswer === 'True' ? styles.correctOption : ''}`}>
              True {question.CorrectAnswer === 'True' && '✓'}
            </div>
            <div className={`${styles.option} ${question.CorrectAnswer === 'False' ? styles.correctOption : ''}`}>
              False {question.CorrectAnswer === 'False' && '✓'}
            </div>
          </div>
        );

      case QuestionType.MultipleChoice:
      case QuestionType.MultipleSelect:
        return (
          <div className={styles.optionsList}>
            {question.OptionA && (
              <div className={`${styles.option} ${question.CorrectAnswer?.includes('A') ? styles.correctOption : ''}`}>
                A. {question.OptionA} {question.CorrectAnswer?.includes('A') && '✓'}
              </div>
            )}
            {question.OptionB && (
              <div className={`${styles.option} ${question.CorrectAnswer?.includes('B') ? styles.correctOption : ''}`}>
                B. {question.OptionB} {question.CorrectAnswer?.includes('B') && '✓'}
              </div>
            )}
            {question.OptionC && (
              <div className={`${styles.option} ${question.CorrectAnswer?.includes('C') ? styles.correctOption : ''}`}>
                C. {question.OptionC} {question.CorrectAnswer?.includes('C') && '✓'}
              </div>
            )}
            {question.OptionD && (
              <div className={`${styles.option} ${question.CorrectAnswer?.includes('D') ? styles.correctOption : ''}`}>
                D. {question.OptionD} {question.CorrectAnswer?.includes('D') && '✓'}
              </div>
            )}
          </div>
        );

      case QuestionType.Matching:
        try {
          const pairs = question.MatchingPairs ? JSON.parse(question.MatchingPairs) : [];
          return (
            <div className={styles.optionsList}>
              <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>Match pairs:</Text>
              {pairs.map((pair: { left: string; right: string }, i: number) => (
                <div key={i} className={styles.option}>
                  {pair.left} → {pair.right}
                </div>
              ))}
            </div>
          );
        } catch {
          return <></>;
        }

      case QuestionType.Ordering:
        try {
          const items = question.OrderingItems ? JSON.parse(question.OrderingItems) : [];
          return (
            <div className={styles.optionsList}>
              <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>Correct order:</Text>
              {items.map((item: string, i: number) => (
                <div key={i} className={styles.option}>
                  {i + 1}. {item}
                </div>
              ))}
            </div>
          );
        } catch {
          return <></>;
        }

      case QuestionType.FillInBlank:
        try {
          const blanks = question.BlankAnswers ? JSON.parse(question.BlankAnswers) : [];
          return (
            <div className={styles.optionsList}>
              <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>Accepted answers:</Text>
              {blanks.map((blank: string, i: number) => (
                <span key={i} className={`${styles.badge} ${styles.badgeSuccess}`} style={{ marginRight: '4px' }}>
                  {blank}
                </span>
              ))}
            </div>
          );
        } catch {
          return <></>;
        }

      case QuestionType.RatingScale:
        return (
          <div className={styles.optionsList}>
            <Text variant="small">
              Scale: {question.ScaleMin || 1} ({question.ScaleLabels ? JSON.parse(question.ScaleLabels).min : 'Min'})
              to {question.ScaleMax || 5} ({question.ScaleLabels ? JSON.parse(question.ScaleLabels).max : 'Max'})
            </Text>
          </div>
        );

      case QuestionType.Essay:
        return (
          <div className={styles.optionsList}>
            <span className={`${styles.badge} ${styles.badgeWarning}`}>
              Long answer - requires manual grading
            </span>
            {question.MaxWordCount && (
              <Text variant="small" style={{ marginLeft: '8px' }}>
                Word limit: {question.MaxWordCount}
              </Text>
            )}
          </div>
        );

      case QuestionType.ShortAnswer:
        return (
          <div className={styles.optionsList}>
            <Text variant="small" className={styles.correctOption}>
              Expected: {question.CorrectAnswer}
            </Text>
          </div>
        );

      default:
        return <></>;
    }
  }

  private renderQuestionPanel(): JSX.Element {
    const { showQuestionPanel, editingQuestion, editingQuestionIndex, saving } = this.state;

    if (!editingQuestion) return <></>;

    const isNew = editingQuestionIndex < 0;

    return (
      <Panel
        isOpen={showQuestionPanel}
        type={PanelType.large}
        onDismiss={() => this.setState({ showQuestionPanel: false, editingQuestion: null })}
        headerText={isNew ? 'Add Question' : 'Edit Question'}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Save Question"
              onClick={() => { void this.handleSaveQuestion(); }}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showQuestionPanel: false, editingQuestion: null })}
            />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <div className={styles.panelContent}>
          <Stack tokens={{ childrenGap: 16 }}>
            {/* Question Type Selection */}
            <Dropdown
              label="Question Type"
              options={questionTypeOptions}
              selectedKey={editingQuestion.QuestionType}
              onChange={(e, option) => this.setState({
                editingQuestion: {
                  ...editingQuestion,
                  QuestionType: option?.key as QuestionType,
                  CorrectAnswer: option?.key === QuestionType.TrueFalse ? 'True' :
                                 option?.key === QuestionType.Essay ? '' : 'A'
                }
              })}
              onRenderOption={(option) => (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <IconButton iconProps={{ iconName: option?.data?.icon }} disabled styles={{ root: { height: 20, width: 20 } }} />
                  <Text>{option?.text}</Text>
                </Stack>
              )}
            />

            {/* Question Text */}
            <TextField
              label="Question Text"
              value={editingQuestion.QuestionText || ''}
              onChange={(e, value) => this.setState({
                editingQuestion: { ...editingQuestion, QuestionText: value || '' }
              })}
              multiline
              rows={3}
              required
              placeholder="Enter your question..."
              description={editingQuestion.QuestionType === QuestionType.FillInBlank
                ? 'Use _____ (underscores) to indicate blank spaces'
                : undefined}
            />

            {/* Question metadata */}
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack.Item grow={1}>
                <Dropdown
                  label="Difficulty"
                  options={difficultyOptions}
                  selectedKey={editingQuestion.DifficultyLevel || DifficultyLevel.Medium}
                  onChange={(e, option) => this.setState({
                    editingQuestion: { ...editingQuestion, DifficultyLevel: option?.key as DifficultyLevel }
                  })}
                />
              </Stack.Item>
              <Stack.Item>
                <TextField
                  label="Points"
                  type="number"
                  value={String(editingQuestion.Points || 10)}
                  onChange={(e, value) => this.setState({
                    editingQuestion: { ...editingQuestion, Points: parseInt(value || '10', 10) }
                  })}
                  min={1}
                  max={100}
                  styles={{ root: { width: 80 } }}
                />
              </Stack.Item>
              <Stack.Item>
                <TextField
                  label="Time Limit (sec)"
                  type="number"
                  value={String(editingQuestion.TimeLimit || 0)}
                  onChange={(e, value) => this.setState({
                    editingQuestion: { ...editingQuestion, TimeLimit: parseInt(value || '0', 10) }
                  })}
                  min={0}
                  max={600}
                  styles={{ root: { width: 100 } }}
                  description="0 = no limit"
                />
              </Stack.Item>
            </Stack>

            <Separator />

            {/* Type-specific editors */}
            {this.renderQuestionTypeEditor()}

            <Separator />

            {/* Feedback and hints */}
            <TextField
              label="Explanation (shown after answering)"
              value={editingQuestion.Explanation || ''}
              onChange={(e, value) => this.setState({
                editingQuestion: { ...editingQuestion, Explanation: value || '' }
              })}
              multiline
              rows={2}
              placeholder="Explain why this is the correct answer..."
            />

            <TextField
              label="Hint (optional)"
              value={editingQuestion.Hint || ''}
              onChange={(e, value) => this.setState({
                editingQuestion: { ...editingQuestion, Hint: value || '' }
              })}
              placeholder="Provide a hint for struggling users..."
            />

            <TextField
              label="Tags (comma-separated)"
              value={editingQuestion.Tags || ''}
              onChange={(e, value) => this.setState({
                editingQuestion: { ...editingQuestion, Tags: value || '' }
              })}
              placeholder="compliance, security, onboarding..."
            />

            <Toggle
              label="Question Active"
              checked={editingQuestion.IsActive !== false}
              onChange={(e, checked) => this.setState({
                editingQuestion: { ...editingQuestion, IsActive: checked || false }
              })}
            />
          </Stack>
        </div>
      </Panel>
    );
  }

  private renderQuestionTypeEditor(): JSX.Element {
    const { editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    switch (editingQuestion.QuestionType) {
      case QuestionType.MultipleChoice:
        return this.renderMultipleChoiceEditor();
      case QuestionType.TrueFalse:
        return this.renderTrueFalseEditor();
      case QuestionType.MultipleSelect:
        return this.renderMultipleSelectEditor();
      case QuestionType.ShortAnswer:
        return this.renderShortAnswerEditor();
      case QuestionType.FillInBlank:
        return this.renderFillInBlankEditor();
      case QuestionType.Matching:
        return this.renderMatchingEditor();
      case QuestionType.Ordering:
        return this.renderOrderingEditor();
      case QuestionType.RatingScale:
        return this.renderRatingScaleEditor();
      case QuestionType.Essay:
        return this.renderEssayEditor();
      case QuestionType.ImageChoice:
        return this.renderImageChoiceEditor();
      case QuestionType.Hotspot:
        return this.renderHotspotEditor();
      default:
        return <></>;
    }
  }

  private renderMultipleChoiceEditor(): JSX.Element {
    const { editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Answer Options</Label>
        <TextField
          prefix="A"
          value={editingQuestion.OptionA || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, OptionA: value || '' }
          })}
          placeholder="Enter option A..."
        />
        <TextField
          prefix="B"
          value={editingQuestion.OptionB || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, OptionB: value || '' }
          })}
          placeholder="Enter option B..."
        />
        <TextField
          prefix="C"
          value={editingQuestion.OptionC || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, OptionC: value || '' }
          })}
          placeholder="Enter option C (optional)..."
        />
        <TextField
          prefix="D"
          value={editingQuestion.OptionD || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, OptionD: value || '' }
          })}
          placeholder="Enter option D (optional)..."
        />

        <div className={styles.correctAnswerSection}>
          <Dropdown
            label="Correct Answer"
            options={[
              { key: 'A', text: 'Option A' },
              { key: 'B', text: 'Option B' },
              { key: 'C', text: 'Option C' },
              { key: 'D', text: 'Option D' }
            ]}
            selectedKey={editingQuestion.CorrectAnswer || 'A'}
            onChange={(e, option) => this.setState({
              editingQuestion: { ...editingQuestion, CorrectAnswer: option?.key as string }
            })}
          />
        </div>
      </Stack>
    );
  }

  private renderTrueFalseEditor(): JSX.Element {
    const { editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <div className={styles.correctAnswerSection}>
        <ChoiceGroup
          label="Correct Answer"
          options={[
            { key: 'True', text: 'True' },
            { key: 'False', text: 'False' }
          ]}
          selectedKey={editingQuestion.CorrectAnswer || 'True'}
          onChange={(e, option) => this.setState({
            editingQuestion: { ...editingQuestion, CorrectAnswer: option?.key || 'True' }
          })}
        />
      </div>
    );
  }

  private renderMultipleSelectEditor(): JSX.Element {
    const { editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    const selectedAnswers = (editingQuestion.CorrectAnswers || editingQuestion.CorrectAnswer || '').split(';').filter(Boolean);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Answer Options (select all correct answers)</Label>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Checkbox
            checked={selectedAnswers.includes('A')}
            onChange={(e, checked) => {
              const answers = new Set(selectedAnswers);
              if (checked) answers.add('A'); else answers.delete('A');
              this.setState({
                editingQuestion: {
                  ...editingQuestion,
                  CorrectAnswer: Array.from(answers).join(';'),
                  CorrectAnswers: Array.from(answers).join(';')
                }
              });
            }}
          />
          <TextField
            prefix="A"
            value={editingQuestion.OptionA || ''}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, OptionA: value || '' }
            })}
            placeholder="Enter option A..."
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Checkbox
            checked={selectedAnswers.includes('B')}
            onChange={(e, checked) => {
              const answers = new Set(selectedAnswers);
              if (checked) answers.add('B'); else answers.delete('B');
              this.setState({
                editingQuestion: {
                  ...editingQuestion,
                  CorrectAnswer: Array.from(answers).join(';'),
                  CorrectAnswers: Array.from(answers).join(';')
                }
              });
            }}
          />
          <TextField
            prefix="B"
            value={editingQuestion.OptionB || ''}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, OptionB: value || '' }
            })}
            placeholder="Enter option B..."
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Checkbox
            checked={selectedAnswers.includes('C')}
            onChange={(e, checked) => {
              const answers = new Set(selectedAnswers);
              if (checked) answers.add('C'); else answers.delete('C');
              this.setState({
                editingQuestion: {
                  ...editingQuestion,
                  CorrectAnswer: Array.from(answers).join(';'),
                  CorrectAnswers: Array.from(answers).join(';')
                }
              });
            }}
          />
          <TextField
            prefix="C"
            value={editingQuestion.OptionC || ''}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, OptionC: value || '' }
            })}
            placeholder="Enter option C (optional)..."
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Checkbox
            checked={selectedAnswers.includes('D')}
            onChange={(e, checked) => {
              const answers = new Set(selectedAnswers);
              if (checked) answers.add('D'); else answers.delete('D');
              this.setState({
                editingQuestion: {
                  ...editingQuestion,
                  CorrectAnswer: Array.from(answers).join(';'),
                  CorrectAnswers: Array.from(answers).join(';')
                }
              });
            }}
          />
          <TextField
            prefix="D"
            value={editingQuestion.OptionD || ''}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, OptionD: value || '' }
            })}
            placeholder="Enter option D (optional)..."
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <MessageBar messageBarType={MessageBarType.info}>
          Check all options that are correct answers. Partial credit can be awarded based on quiz settings.
        </MessageBar>
      </Stack>
    );
  }

  private renderShortAnswerEditor(): JSX.Element {
    const { editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField
          label="Correct Answer"
          value={editingQuestion.CorrectAnswer || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, CorrectAnswer: value || '' }
          })}
          placeholder="Enter the expected answer..."
        />
        <Toggle
          label="Case Sensitive"
          checked={editingQuestion.CaseSensitive || false}
          onChange={(e, checked) => this.setState({
            editingQuestion: { ...editingQuestion, CaseSensitive: checked }
          })}
        />
        <MessageBar messageBarType={MessageBarType.info}>
          The user's answer will be compared to this. Consider enabling case-insensitive matching.
        </MessageBar>
      </Stack>
    );
  }

  private renderFillInBlankEditor(): JSX.Element {
    const { blankAnswers } = this.state;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Accepted Answers for Each Blank</Label>
        <MessageBar messageBarType={MessageBarType.info}>
          Use _____ (underscores) in the question text to indicate blanks. Add acceptable answers for each blank below.
        </MessageBar>

        {blankAnswers.map((answer, index) => (
          <Stack key={index} horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <TextField
              prefix={`Blank ${index + 1}`}
              value={answer}
              onChange={(e, value) => {
                const updated = [...blankAnswers];
                updated[index] = value || '';
                this.setState({ blankAnswers: updated });
              }}
              placeholder="Enter accepted answer..."
              styles={{ root: { flex: 1 } }}
            />
            {blankAnswers.length > 1 && (
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => {
                  const updated = blankAnswers.filter((_, i) => i !== index);
                  this.setState({ blankAnswers: updated });
                }}
              />
            )}
          </Stack>
        ))}

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Another Blank"
          onClick={() => this.setState({ blankAnswers: [...blankAnswers, ''] })}
        />
      </Stack>
    );
  }

  private renderMatchingEditor(): JSX.Element {
    const { matchingItems } = this.state;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Matching Pairs</Label>
        <MessageBar messageBarType={MessageBarType.info}>
          Create pairs of items that users need to match. The items will be shuffled when displayed.
        </MessageBar>

        {matchingItems.map((pair, index) => (
          <div key={index} className={styles.matchingPairRow}>
            <TextField
              value={pair.left}
              onChange={(e, value) => {
                const updated = [...matchingItems];
                updated[index] = { ...updated[index], left: value || '' };
                this.setState({ matchingItems: updated });
              }}
              placeholder="Left item..."
              styles={{ root: { flex: 1 } }}
            />
            <IconButton iconProps={{ iconName: 'Forward' }} disabled />
            <TextField
              value={pair.right}
              onChange={(e, value) => {
                const updated = [...matchingItems];
                updated[index] = { ...updated[index], right: value || '' };
                this.setState({ matchingItems: updated });
              }}
              placeholder="Matching item..."
              styles={{ root: { flex: 1 } }}
            />
            {matchingItems.length > 1 && (
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => {
                  const updated = matchingItems.filter((_, i) => i !== index);
                  this.setState({ matchingItems: updated });
                }}
              />
            )}
          </div>
        ))}

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Matching Pair"
          onClick={() => this.setState({ matchingItems: [...matchingItems, { left: '', right: '' }] })}
        />
      </Stack>
    );
  }

  private renderOrderingEditor(): JSX.Element {
    const { orderingItems } = this.state;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Items to Order (in correct sequence)</Label>
        <MessageBar messageBarType={MessageBarType.info}>
          Enter items in the correct order. They will be shuffled for the user to reorder.
        </MessageBar>

        {orderingItems.map((item, index) => (
          <div key={index} className={styles.orderingItem}>
            <span className={styles.dragHandle}>
              <IconButton iconProps={{ iconName: 'GripperBarHorizontal' }} disabled />
            </span>
            <Text style={{ width: '24px' }}>{index + 1}.</Text>
            <TextField
              value={item}
              onChange={(e, value) => {
                const updated = [...orderingItems];
                updated[index] = value || '';
                this.setState({ orderingItems: updated });
              }}
              placeholder={`Step ${index + 1}...`}
              styles={{ root: { flex: 1 } }}
            />
            {orderingItems.length > 1 && (
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => {
                  const updated = orderingItems.filter((_, i) => i !== index);
                  this.setState({ orderingItems: updated });
                }}
              />
            )}
          </div>
        ))}

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Item"
          onClick={() => this.setState({ orderingItems: [...orderingItems, ''] })}
        />
      </Stack>
    );
  }

  private renderRatingScaleEditor(): JSX.Element {
    const { ratingMin, ratingMax, ratingLabels, editingQuestion } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Label>Rating Scale Configuration</Label>

        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <TextField
            label="Minimum Value"
            type="number"
            value={String(ratingMin)}
            onChange={(e, value) => this.setState({ ratingMin: parseInt(value || '1', 10) })}
            min={0}
            max={10}
            styles={{ root: { width: 100 } }}
          />
          <TextField
            label="Maximum Value"
            type="number"
            value={String(ratingMax)}
            onChange={(e, value) => this.setState({ ratingMax: parseInt(value || '5', 10) })}
            min={1}
            max={10}
            styles={{ root: { width: 100 } }}
          />
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <TextField
            label="Minimum Label"
            value={ratingLabels.min}
            onChange={(e, value) => this.setState({
              ratingLabels: { ...ratingLabels, min: value || '' }
            })}
            placeholder="e.g., Poor, Disagree, Never"
          />
          <TextField
            label="Maximum Label"
            value={ratingLabels.max}
            onChange={(e, value) => this.setState({
              ratingLabels: { ...ratingLabels, max: value || '' }
            })}
            placeholder="e.g., Excellent, Agree, Always"
          />
        </Stack>

        <MessageBar messageBarType={MessageBarType.info}>
          Rating scale questions are typically used for surveys and feedback. No correct answer is required.
        </MessageBar>
      </Stack>
    );
  }

  private renderEssayEditor(): JSX.Element {
    const { editingQuestion, gradingRubric } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <TextField
            label="Max Word Count"
            type="number"
            value={String(editingQuestion.MaxWordCount || 500)}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, MaxWordCount: parseInt(value || '500', 10) }
            })}
            min={50}
            max={5000}
            styles={{ root: { width: 120 } }}
          />
          <TextField
            label="Min Word Count"
            type="number"
            value={String(editingQuestion.MinWordCount || 50)}
            onChange={(e, value) => this.setState({
              editingQuestion: { ...editingQuestion, MinWordCount: parseInt(value || '50', 10) }
            })}
            min={10}
            max={1000}
            styles={{ root: { width: 120 } }}
          />
        </Stack>

        <MessageBar messageBarType={MessageBarType.warning}>
          Essay questions require manual grading. Add a rubric below to guide the grading process.
        </MessageBar>

        <Label>Grading Rubric</Label>
        {gradingRubric.map((item, index) => (
          <div key={index} className={styles.rubricRow}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField
                  label="Criteria"
                  value={item.criteria}
                  onChange={(e, value) => {
                    const updated = [...gradingRubric];
                    updated[index] = { ...updated[index], criteria: value || '' };
                    this.setState({ gradingRubric: updated });
                  }}
                  placeholder="e.g., Content Quality"
                  styles={{ root: { flex: 2 } }}
                />
                <TextField
                  label="Max Points"
                  type="number"
                  value={String(item.maxPoints)}
                  onChange={(e, value) => {
                    const updated = [...gradingRubric];
                    updated[index] = { ...updated[index], maxPoints: parseInt(value || '0', 10) };
                    this.setState({ gradingRubric: updated });
                  }}
                  styles={{ root: { width: 80 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  onClick={() => {
                    const updated = gradingRubric.filter((_, i) => i !== index);
                    this.setState({ gradingRubric: updated });
                  }}
                  style={{ marginTop: '28px' }}
                />
              </Stack>
              <TextField
                label="Description"
                value={item.description}
                onChange={(e, value) => {
                  const updated = [...gradingRubric];
                  updated[index] = { ...updated[index], description: value || '' };
                  this.setState({ gradingRubric: updated });
                }}
                placeholder="Describe expectations for full marks..."
              />
            </Stack>
          </div>
        ))}

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Rubric Criteria"
          onClick={() => this.setState({
            gradingRubric: [...gradingRubric, { criteria: '', maxPoints: 10, description: '' }]
          })}
        />
      </Stack>
    );
  }

  private renderImageChoiceEditor(): JSX.Element {
    const { imageChoices } = this.state;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Label>Image Answer Options</Label>
        <MessageBar messageBarType={MessageBarType.info}>
          Add images with labels. Users will select the correct image(s).
        </MessageBar>

        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          {imageChoices.map((choice, index) => (
            <div
              key={index}
              className={`${styles.imageChoiceCard} ${choice.isCorrect ? styles.imageChoiceSelected : ''}`}
            >
              <TextField
                label="Image URL"
                value={choice.imageUrl}
                onChange={(e, value) => {
                  const updated = [...imageChoices];
                  updated[index] = { ...updated[index], imageUrl: value || '' };
                  this.setState({ imageChoices: updated });
                }}
                placeholder="https://..."
              />
              <TextField
                label="Label"
                value={choice.label}
                onChange={(e, value) => {
                  const updated = [...imageChoices];
                  updated[index] = { ...updated[index], label: value || '' };
                  this.setState({ imageChoices: updated });
                }}
              />
              <Checkbox
                label="Correct Answer"
                checked={choice.isCorrect}
                onChange={(e, checked) => {
                  const updated = [...imageChoices];
                  updated[index] = { ...updated[index], isCorrect: checked || false };
                  this.setState({ imageChoices: updated });
                }}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => {
                  const updated = imageChoices.filter((_, i) => i !== index);
                  this.setState({ imageChoices: updated });
                }}
              />
            </div>
          ))}
        </Stack>

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Image Option"
          onClick={() => this.setState({
            imageChoices: [...imageChoices, { imageUrl: '', label: '', isCorrect: false }]
          })}
        />
      </Stack>
    );
  }

  private renderHotspotEditor(): JSX.Element {
    const { editingQuestion, hotspotRegions } = this.state;
    if (!editingQuestion) return <></>;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <TextField
          label="Base Image URL"
          value={editingQuestion.QuestionImage || ''}
          onChange={(e, value) => this.setState({
            editingQuestion: { ...editingQuestion, QuestionImage: value || '' }
          })}
          placeholder="https://..."
        />

        <MessageBar messageBarType={MessageBarType.info}>
          Define clickable regions on the image. Users must click within these regions to answer correctly.
        </MessageBar>

        <Label>Hotspot Regions</Label>
        {hotspotRegions.map((region, index) => (
          <div key={index} className={styles.rubricRow}>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="X Position (%)"
                type="number"
                value={String(region.x)}
                onChange={(e, value) => {
                  const updated = [...hotspotRegions];
                  updated[index] = { ...updated[index], x: parseInt(value || '0', 10) };
                  this.setState({ hotspotRegions: updated });
                }}
                min={0}
                max={100}
                styles={{ root: { width: 80 } }}
              />
              <TextField
                label="Y Position (%)"
                type="number"
                value={String(region.y)}
                onChange={(e, value) => {
                  const updated = [...hotspotRegions];
                  updated[index] = { ...updated[index], y: parseInt(value || '0', 10) };
                  this.setState({ hotspotRegions: updated });
                }}
                min={0}
                max={100}
                styles={{ root: { width: 80 } }}
              />
              <TextField
                label="Width (%)"
                type="number"
                value={String(region.width)}
                onChange={(e, value) => {
                  const updated = [...hotspotRegions];
                  updated[index] = { ...updated[index], width: parseInt(value || '10', 10) };
                  this.setState({ hotspotRegions: updated });
                }}
                min={1}
                max={100}
                styles={{ root: { width: 80 } }}
              />
              <TextField
                label="Height (%)"
                type="number"
                value={String(region.height)}
                onChange={(e, value) => {
                  const updated = [...hotspotRegions];
                  updated[index] = { ...updated[index], height: parseInt(value || '10', 10) };
                  this.setState({ hotspotRegions: updated });
                }}
                min={1}
                max={100}
                styles={{ root: { width: 80 } }}
              />
              <TextField
                label="Label"
                value={region.label}
                onChange={(e, value) => {
                  const updated = [...hotspotRegions];
                  updated[index] = { ...updated[index], label: value || '' };
                  this.setState({ hotspotRegions: updated });
                }}
                styles={{ root: { width: 120 } }}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => {
                  const updated = hotspotRegions.filter((_, i) => i !== index);
                  this.setState({ hotspotRegions: updated });
                }}
                style={{ marginTop: '28px' }}
              />
            </Stack>
          </div>
        ))}

        <ActionButton
          iconProps={{ iconName: 'Add' }}
          text="Add Hotspot Region"
          onClick={() => this.setState({
            hotspotRegions: [...hotspotRegions, { x: 10, y: 10, width: 20, height: 20, label: '' }]
          })}
        />
      </Stack>
    );
  }

  private renderDeleteDialog(): JSX.Element {
    const { showDeleteDialog, saving } = this.state;

    return (
      <Dialog
        hidden={!showDeleteDialog}
        onDismiss={() => this.setState({ showDeleteDialog: false, questionToDelete: null })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Question',
          subText: 'Are you sure you want to delete this question? This action cannot be undone.'
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text="Delete"
            onClick={() => { void this.handleDeleteQuestion(); }}
            disabled={saving}
            styles={{ root: { backgroundColor: '#a80000' } }}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => this.setState({ showDeleteDialog: false, questionToDelete: null })}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  private renderImportDialog(): JSX.Element {
    const { showImportDialog, importFormat, saving } = this.state;

    return (
      <Dialog
        hidden={!showImportDialog}
        onDismiss={() => this.setState({ showImportDialog: false, importFile: null })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Import Questions',
          subText: 'Import questions from a JSON or CSV file.'
        }}
        minWidth={450}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <ChoiceGroup
            label="Import Format"
            options={[
              { key: 'json', text: 'JSON (full question data)' },
              { key: 'csv', text: 'CSV (basic question data)' }
            ]}
            selectedKey={importFormat}
            onChange={(e, option) => this.setState({ importFormat: option?.key as 'json' | 'csv' })}
          />

          <input
            type="file"
            ref={this.fileInputRef}
            accept={importFormat === 'json' ? '.json' : '.csv'}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) {
                this.setState({ importFile: file });
              }
            }}
            style={{ display: 'none' }}
          />

          <DefaultButton
            text="Select File"
            iconProps={{ iconName: 'Attach' }}
            onClick={() => this.fileInputRef.current?.click()}
          />

          {this.state.importFile && (
            <MessageBar messageBarType={MessageBarType.success}>
              Selected: {this.state.importFile.name}
            </MessageBar>
          )}
        </Stack>

        <DialogFooter>
          <PrimaryButton
            text="Import"
            onClick={() => { void this.handleImportQuestions(); }}
            disabled={saving || !this.state.importFile}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => this.setState({ showImportDialog: false, importFile: null })}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  private renderExportDialog(): JSX.Element {
    const { showExportDialog, exportFormat, questions } = this.state;

    return (
      <Dialog
        hidden={!showExportDialog}
        onDismiss={() => this.setState({ showExportDialog: false })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Export Questions',
          subText: `Export ${questions.length} questions to a file.`
        }}
        minWidth={400}
      >
        <ChoiceGroup
          label="Export Format"
          options={[
            { key: 'json', text: 'JSON (full question data, can be re-imported)' },
            { key: 'csv', text: 'CSV (basic data, Excel compatible)' }
          ]}
          selectedKey={exportFormat}
          onChange={(e, option) => this.setState({ exportFormat: option?.key as 'json' | 'csv' })}
        />

        <DialogFooter>
          <PrimaryButton
            text="Export"
            onClick={() => { void this.handleExportQuestions(); }}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => this.setState({ showExportDialog: false })}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  private renderStatisticsPanel(): JSX.Element {
    const { showStatisticsPanel, quizStatistics, title } = this.state;

    return (
      <Panel
        isOpen={showStatisticsPanel}
        type={PanelType.medium}
        onDismiss={() => this.setState({ showStatisticsPanel: false })}
        headerText={`Statistics: ${title}`}
        closeButtonAriaLabel="Close"
      >
        {!quizStatistics ? (
          <Spinner label="Loading statistics..." />
        ) : (
          <Stack tokens={{ childrenGap: 20 }} style={{ padding: '20px 0' }}>
            {/* Summary Cards */}
            <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
              <div className={styles.statisticsCard}>
                <Text className={styles.statValue}>{quizStatistics.totalAttempts}</Text>
                <Text className={styles.statLabel}>Total Attempts</Text>
              </div>
              <div className={styles.statisticsCard}>
                <Text className={styles.statValue}>{quizStatistics.passRate.toFixed(1)}%</Text>
                <Text className={styles.statLabel}>Pass Rate</Text>
              </div>
              <div className={styles.statisticsCard}>
                <Text className={styles.statValue}>{quizStatistics.averageScore.toFixed(1)}%</Text>
                <Text className={styles.statLabel}>Average Score</Text>
              </div>
              <div className={styles.statisticsCard}>
                <Text className={styles.statValue}>{Math.round(quizStatistics.averageTimeSpent / 60)} min</Text>
                <Text className={styles.statLabel}>Avg Time</Text>
              </div>
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="large">Score Distribution</Text>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {quizStatistics.scoreDistribution.map((dist) => (
                  <div key={dist.range} style={{ textAlign: 'center' }}>
                    <Text style={{ fontWeight: 600 }}>{dist.count}</Text>
                    <Text variant="small">{dist.range}</Text>
                  </div>
                ))}
              </Stack>
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="large">Completion Stats</Text>
              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Text>Completed: {quizStatistics.completionRate.toFixed(1)}%</Text>
                <Text>Highest: {quizStatistics.highestScore}%</Text>
                <Text>Lowest: {quizStatistics.lowestScore}%</Text>
              </Stack>
            </Stack>

            {quizStatistics.questionAnalytics && quizStatistics.questionAnalytics.length > 0 && (
              <>
                <Separator />
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text variant="large">Question Performance</Text>
                  {quizStatistics.questionAnalytics.slice(0, 5).map((qa, index) => (
                    <Stack key={index} tokens={{ childrenGap: 4 }} style={{ padding: '8px', backgroundColor: '#f8f8f8', borderRadius: '4px' }}>
                      <Text style={{ fontWeight: 500 }}>Q{index + 1}: {qa.questionText?.substring(0, 50)}...</Text>
                      <Stack horizontal tokens={{ childrenGap: 16 }}>
                        <Text variant="small">Correct: {(qa.correctRate * 100).toFixed(1)}%</Text>
                        <Text variant="small">Avg Time: {qa.averageTime}s</Text>
                        <Text variant="small">Difficulty: {(qa.difficultyIndex * 100).toFixed(0)}%</Text>
                      </Stack>
                    </Stack>
                  ))}
                </Stack>
              </>
            )}
          </Stack>
        )}
      </Panel>
    );
  }

  // ============================================================================
  // AI QUESTION BUILDER
  // ============================================================================

  private handleAiGenerate = async (): Promise<void> => {
    const { aiFunctionUrl, aiQuestionCount, aiDifficulty, aiIncludeExcerpts, aiSelectedTypes, policyId } = this.state;

    if (!aiFunctionUrl.trim()) {
      this.setState({ aiError: 'Please enter the Azure Function URL.' });
      return;
    }

    this.setState({ aiGenerating: true, aiError: null, aiGeneratedQuestions: [] });

    try {
      // Find the policy to get its document URL and metadata
      const policy = this.state.policies.find(p => p.Id === policyId);
      const policyTitle = policy?.PolicyName || 'Policy';
      const policyCategory = policy?.PolicyCategory || 'General';

      // Fetch the document text from SharePoint if possible
      let policyText = '';
      const docUrl = policy?.DocumentURL;
      if (docUrl && typeof docUrl === 'string') {
        try {
          // Attempt to fetch document content via SharePoint REST API
          const fileResponse = await this.props.sp.web.getFileByServerRelativePath(docUrl).getText();
          policyText = fileResponse;
        } catch {
          console.warn('Could not extract document text via SP — will pass URL to Azure Function');
        }
      }

      const requestBody: Record<string, unknown> = {
        questionCount: aiQuestionCount,
        difficultyLevel: aiDifficulty,
        questionTypes: aiSelectedTypes,
        includeExcerpts: aiIncludeExcerpts,
        policyTitle,
        policyCategory
      };

      if (policyText) {
        requestBody.policyText = policyText;
      } else if (docUrl) {
        requestBody.policyDocumentUrl = docUrl;
      } else {
        this.setState({ aiGenerating: false, aiError: 'No policy document URL found. Please link a document to this policy first.' });
        return;
      }

      const response = await fetch(aiFunctionUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: `HTTP ${response.status}` })) as { error: string };
        throw new Error(errorData.error || `HTTP ${response.status}`);
      }

      const data = await response.json() as { questions: IQuizQuestion[]; metadata: Record<string, unknown> };

      this.setState({
        aiGenerating: false,
        aiGeneratedQuestions: data.questions || [],
        success: `AI generated ${(data.questions || []).length} questions in ${data.metadata?.generationTimeMs || 0}ms`
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      this.setState({ aiGenerating: false, aiError: `AI generation failed: ${message}` });
    }
  };

  private handleImportAiQuestions = async (): Promise<void> => {
    const { aiGeneratedQuestions, quizId } = this.state;
    if (!quizId || aiGeneratedQuestions.length === 0) return;

    this.setState({ saving: true });

    try {
      for (const question of aiGeneratedQuestions) {
        await this.quizService.createQuestion({
          ...question,
          QuizId: quizId,
          IsActive: true,
          IsRequired: false,
          QuestionOrder: this.state.questions.length + 1
        });
      }

      const updatedQuestions = await this.quizService.getQuizQuestions(quizId);
      this.setState({
        questions: updatedQuestions,
        saving: false,
        showAiPanel: false,
        aiGeneratedQuestions: [],
        success: `Successfully imported ${aiGeneratedQuestions.length} AI-generated questions!`
      });
    } catch (err) {
      console.error('Failed to import AI questions:', err);
      this.setState({ saving: false, error: 'Failed to import AI-generated questions.' });
    }
  };

  private renderAiPanel(): JSX.Element {
    const {
      showAiPanel, aiQuestionCount, aiDifficulty, aiIncludeExcerpts,
      aiSelectedTypes, aiGenerating, aiError, aiGeneratedQuestions, aiFunctionUrl
    } = this.state;

    const difficultyOptions: IDropdownOption[] = [
      { key: 'Easy', text: 'Easy — Basic facts and definitions' },
      { key: 'Medium', text: 'Medium — Application scenarios' },
      { key: 'Hard', text: 'Hard — Analysis and judgment' },
      { key: 'Expert', text: 'Expert — Complex scenarios and edge cases' }
    ];

    const typeOptions = [
      'Multiple Choice', 'True/False', 'Multiple Select', 'Short Answer',
      'Fill in the Blank', 'Matching', 'Ordering', 'Essay'
    ];

    return (
      <Panel
        isOpen={showAiPanel}
        onDismiss={() => this.setState({ showAiPanel: false })}
        headerText="AI Question Builder"
        type={PanelType.medium}
      >
        <Stack tokens={{ childrenGap: 16, padding: '16px 0' }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Generate quiz questions automatically from your policy document using Azure OpenAI GPT-4.
            The AI will analyze the document and create questions aligned with the selected difficulty and types.
          </MessageBar>

          {/* Azure Function URL */}
          <TextField
            label="Azure Function URL"
            value={aiFunctionUrl}
            onChange={(_e, v) => this.setState({ aiFunctionUrl: v || '' })}
            placeholder="https://your-func.azurewebsites.net/api/generate-quiz-questions?code=..."
            description="Enter your deployed Azure Function endpoint URL with the function key"
            required
          />

          <Separator />

          {/* Question Count */}
          <Slider
            label={`Number of Questions: ${aiQuestionCount}`}
            min={1}
            max={50}
            step={1}
            value={aiQuestionCount}
            showValue={false}
            onChange={(v) => this.setState({ aiQuestionCount: v })}
          />

          {/* Difficulty */}
          <Dropdown
            label="Difficulty Level"
            selectedKey={aiDifficulty}
            options={difficultyOptions}
            onChange={(_e, opt) => opt && this.setState({ aiDifficulty: opt.key as string })}
          />

          {/* Question Types */}
          <Label>Question Types to Generate</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {typeOptions.map(type => (
              <Checkbox
                key={type}
                label={type}
                checked={aiSelectedTypes.includes(type)}
                onChange={(_e, checked) => {
                  const updated = checked
                    ? [...aiSelectedTypes, type]
                    : aiSelectedTypes.filter(t => t !== type);
                  this.setState({ aiSelectedTypes: updated });
                }}
              />
            ))}
          </Stack>

          {/* Include Excerpts */}
          <Toggle
            label="Include Document Excerpts"
            checked={aiIncludeExcerpts}
            onChange={(_e, checked) => this.setState({ aiIncludeExcerpts: !!checked })}
            onText="Yes — AI will include relevant policy passages"
            offText="No — Questions only"
          />

          <Separator />

          {/* Error */}
          {aiError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ aiError: null })}>
              {aiError}
            </MessageBar>
          )}

          {/* Generate Button */}
          <PrimaryButton
            text={aiGenerating ? 'Generating...' : 'Generate Questions'}
            iconProps={{ iconName: 'Robot' }}
            onClick={this.handleAiGenerate}
            disabled={aiGenerating || !aiFunctionUrl.trim()}
          />

          {aiGenerating && (
            <Stack horizontalAlign="center" tokens={{ padding: 16 }}>
              <Spinner size={SpinnerSize.large} label="AI is analyzing the policy document and generating questions..." />
            </Stack>
          )}

          {/* Preview Generated Questions */}
          {aiGeneratedQuestions.length > 0 && (
            <>
              <Separator />
              <Text variant="large" style={{ fontWeight: 600 }}>
                Generated Questions ({aiGeneratedQuestions.length})
              </Text>

              <div style={{ maxHeight: 400, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 12 }}>
                {aiGeneratedQuestions.map((q, i) => (
                  <div key={i} style={{ padding: '12px 0', borderBottom: i < aiGeneratedQuestions.length - 1 ? '1px solid #edebe9' : 'none' }}>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <span style={{ fontWeight: 700, color: '#0d9488', minWidth: 24 }}>{i + 1}.</span>
                      <span style={{ backgroundColor: '#e8f5e9', padding: '2px 8px', borderRadius: 4, fontSize: 11 }}>
                        {q.QuestionType}
                      </span>
                      <span style={{ backgroundColor: '#fff3e0', padding: '2px 8px', borderRadius: 4, fontSize: 11 }}>
                        {q.DifficultyLevel}
                      </span>
                      <span style={{ fontSize: 11, color: '#605e5c' }}>{q.Points} pts</span>
                    </Stack>
                    <Text style={{ marginTop: 4, display: 'block' }}>{q.QuestionText}</Text>
                    {q.Explanation && (
                      <Text variant="small" style={{ color: '#605e5c', marginTop: 4, display: 'block', fontStyle: 'italic' }}>
                        {q.Explanation.substring(0, 120)}...
                      </Text>
                    )}
                  </div>
                ))}
              </div>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <PrimaryButton
                  text={`Import All ${aiGeneratedQuestions.length} Questions`}
                  iconProps={{ iconName: 'Accept' }}
                  onClick={this.handleImportAiQuestions}
                />
                <DefaultButton
                  text="Discard"
                  iconProps={{ iconName: 'Delete' }}
                  onClick={() => this.setState({ aiGeneratedQuestions: [] })}
                />
              </Stack>
            </>
          )}
        </Stack>
      </Panel>
    );
  }

  public render(): React.ReactElement<IQuizBuilderProps> {
    const { loading, saving, error, success, quizId, activeTab } = this.state;

    if (loading) {
      return (
        <div className={styles.container}>
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Spinner size={SpinnerSize.large} label="Loading quiz builder..." />
          </Stack>
        </div>
      );
    }

    return (
      <div className={styles.container}>
        <div className={styles.header}>
          <Text className={styles.title}>
            {quizId ? 'Edit Quiz' : 'Create New Quiz'}
          </Text>
          {quizId && (
            <span className={`${styles.badge} ${styles.badgeInfo}`}>
              ID: {quizId}
            </span>
          )}
        </div>

        {this.renderCommandBar()}

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ error: null })}
            dismissButtonAriaLabel="Close"
          >
            {error}
          </MessageBar>
        )}

        {success && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => this.setState({ success: null })}
            dismissButtonAriaLabel="Close"
          >
            {success}
          </MessageBar>
        )}

        {saving && (
          <MessageBar messageBarType={MessageBarType.info}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Spinner size={SpinnerSize.small} />
              <Text>Saving...</Text>
            </Stack>
          </MessageBar>
        )}

        <Pivot
          selectedKey={activeTab}
          onLinkClick={(item) => this.setState({ activeTab: item?.props.itemKey || 'settings' })}
        >
          <PivotItem headerText="Settings" itemKey="settings" itemIcon="Settings">
            <div className={styles.tabContent}>
              {this.renderQuizSettings()}
            </div>
          </PivotItem>

          <PivotItem headerText="Questions" itemKey="questions" itemIcon="ClipboardList" itemCount={this.state.questions.length}>
            <div className={styles.tabContent}>
              {this.renderQuestions()}
            </div>
          </PivotItem>

          <PivotItem headerText="Advanced" itemKey="advanced" itemIcon="DeveloperTools">
            <div className={styles.tabContent}>
              {this.renderAdvancedSettings()}
            </div>
          </PivotItem>

          <PivotItem headerText="Certificate" itemKey="certificate" itemIcon="Certificate">
            <div className={styles.tabContent}>
              {this.renderCertificateSettings()}
            </div>
          </PivotItem>
        </Pivot>

        {this.renderQuestionPanel()}
        {this.renderDeleteDialog()}
        {this.renderImportDialog()}
        {this.renderExportDialog()}
        {this.renderStatisticsPanel()}
        {this.renderAiPanel()}
      </div>
    );
  }
}

export default QuizBuilder;
