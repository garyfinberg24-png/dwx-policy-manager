import * as React from 'react';
import styles from './QuizTaker.module.scss';
import {
  PrimaryButton,
  DefaultButton,
  Stack,
  Text,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  IconButton,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  TextField,
  Slider,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import {
  QuizService,
  IQuiz,
  IQuizQuestion,
  IQuizAnswer,
  IQuizResult,
  QuestionType
} from '../../services/QuizService';
import { SPFI } from '@pnp/sp';

// ============================================================================
// INTERFACES
// ============================================================================

export interface IQuizTakerUser {
  Id: number;
  Title?: string;
  DisplayName?: string;
  EMail?: string;
  Email?: string;
}

export interface IQuizTakerProps {
  sp: SPFI;
  quizId: number;
  policyId: number;
  userId: IQuizTakerUser | number;
  onComplete?: (result: IQuizResult) => void;
  onCancel?: () => void;
}

/** Extended answer map supporting complex answer types per question type */
type AnswerValue =
  | string                                        // Multiple Choice, True/False, Short Answer, Image Choice, Essay
  | string[]                                       // Multiple Select, Fill in Blank (array of blank values), Ordering (array of item ids)
  | { left: string; right: string }[]              // Matching
  | { x: number; y: number }                       // Hotspot
  | number;                                        // Rating Scale

export interface IQuizTakerState {
  quiz: IQuiz | null;
  questions: IQuizQuestion[];
  currentQuestionIndex: number;
  answers: Map<number, AnswerValue>;
  timeRemaining: number;
  attemptId: number | null;
  isSubmitting: boolean;
  isLoading: boolean;
  result: IQuizResult | null;
  error: string | null;
  showResults: boolean;
  startTime: Date;
}

// ============================================================================
// COMPONENT
// ============================================================================

export class QuizTaker extends React.Component<IQuizTakerProps, IQuizTakerState> {
  private quizService: QuizService;
  private timerInterval: ReturnType<typeof setInterval> | null = null;

  constructor(props: IQuizTakerProps) {
    super(props);

    this.quizService = new QuizService(props.sp);
    this.state = {
      quiz: null,
      questions: [],
      currentQuestionIndex: 0,
      answers: new Map(),
      timeRemaining: 0,
      attemptId: null,
      isSubmitting: false,
      isLoading: true,
      result: null,
      error: null,
      showResults: false,
      startTime: new Date()
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadQuiz();
  }

  public componentWillUnmount(): void {
    this.clearTimer();
  }

  // --------------------------------------------------------------------------
  // Timer helpers
  // --------------------------------------------------------------------------

  private clearTimer(): void {
    if (this.timerInterval !== null) {
      clearInterval(this.timerInterval);
      this.timerInterval = null;
    }
  }

  private startTimer(): void {
    this.clearTimer(); // prevent double timers
    this.timerInterval = setInterval(() => {
      this.setState(prev => {
        const newTimeRemaining = prev.timeRemaining - 1;

        if (newTimeRemaining <= 0) {
          this.clearTimer();
          this.handleTimeExpired();
          return { ...prev, timeRemaining: 0 };
        }

        return { ...prev, timeRemaining: newTimeRemaining };
      });
    }, 1000);
  }

  // --------------------------------------------------------------------------
  // Quiz lifecycle
  // --------------------------------------------------------------------------

  private async loadQuiz(): Promise<void> {
    this.setState({ isLoading: true, error: null });

    try {
      const quiz = await this.quizService.getQuizById(this.props.quizId);
      if (!quiz) {
        this.setState({ error: "Quiz not found", isLoading: false });
        return;
      }

      // Resolve userId
      const resolvedUserId = typeof this.props.userId === 'number'
        ? this.props.userId
        : this.props.userId.Id;

      // Check eligibility
      const eligibility = await this.quizService.canUserTakeQuiz(this.props.quizId, resolvedUserId);
      if (!eligibility.canTake) {
        this.setState({ error: eligibility.reason || "Cannot take quiz", isLoading: false });
        return;
      }

      // Get questions
      const questions = await this.quizService.getQuizQuestions(this.props.quizId, { randomize: quiz.RandomizeQuestions });

      // Start attempt
      const user = typeof this.props.userId === 'number'
        ? { Id: this.props.userId, Title: 'Unknown User', EMail: '' }
        : this.props.userId;
      const userName = user.Title || user.DisplayName || 'Unknown User';
      const userEmail = user.EMail || user.Email || '';
      const attempt = await this.quizService.startQuizAttempt(
        this.props.quizId,
        this.props.policyId,
        user.Id,
        userName,
        userEmail
      );

      if (!attempt) {
        this.setState({ error: "Failed to start quiz attempt", isLoading: false });
        return;
      }

      this.setState({
        quiz,
        questions,
        attemptId: attempt.Id,
        timeRemaining: quiz.TimeLimit * 60,
        startTime: new Date(),
        isLoading: false
      });

      this.startTimer();
    } catch (error) {
      console.error("Failed to load quiz:", error);
      this.setState({ error: "Failed to load quiz", isLoading: false });
    }
  }

  /** Reset state and re-load quiz (used for retake instead of window.location.reload) */
  private async retakeQuiz(): Promise<void> {
    this.clearTimer();
    this.setState({
      quiz: null,
      questions: [],
      currentQuestionIndex: 0,
      answers: new Map(),
      timeRemaining: 0,
      attemptId: null,
      isSubmitting: false,
      isLoading: true,
      result: null,
      error: null,
      showResults: false,
      startTime: new Date()
    });
    await this.loadQuiz();
  }

  private handleTimeExpired(): void {
    this.setState({ error: "Time expired! Quiz will be submitted automatically." });
    setTimeout(() => {
      this.handleSubmit().catch(err => console.error("Auto-submit failed:", err));
    }, 2000);
  }

  // --------------------------------------------------------------------------
  // Answer handling
  // --------------------------------------------------------------------------

  private handleAnswerChange = (questionId: number, answer: AnswerValue): void => {
    this.setState(prev => {
      const newAnswers = new Map(prev.answers);
      newAnswers.set(questionId, answer);
      return { answers: newAnswers };
    });
  };

  // --------------------------------------------------------------------------
  // Navigation
  // --------------------------------------------------------------------------

  private handleNext = (): void => {
    this.setState(prev => ({
      currentQuestionIndex: Math.min(prev.currentQuestionIndex + 1, prev.questions.length - 1)
    }));
  };

  private handlePrevious = (): void => {
    this.setState(prev => ({
      currentQuestionIndex: Math.max(prev.currentQuestionIndex - 1, 0)
    }));
  };

  private handleJumpToQuestion = (index: number): void => {
    this.setState({ currentQuestionIndex: index });
  };

  private handleCancel = (): void => {
    // Clear timer before calling onCancel to prevent leak
    this.clearTimer();

    // Abandon attempt if possible
    if (this.state.attemptId) {
      this.quizService.abandonQuizAttempt(this.state.attemptId).catch(err => {
        console.warn("Failed to abandon quiz attempt:", err);
      });
    }

    if (this.props.onCancel) {
      this.props.onCancel();
    }
  };

  // --------------------------------------------------------------------------
  // Submission
  // --------------------------------------------------------------------------

  private async handleSubmit(): Promise<void> {
    const { quiz, questions, answers, attemptId } = this.state;

    if (!quiz || !attemptId) return;

    // Check if all questions are answered
    const unansweredCount = questions.filter(q => !answers.has(q.Id)).length;
    if (unansweredCount > 0) {
      if (!window.confirm(`You have ${unansweredCount} unanswered questions. Submit anyway?`)) {
        return;
      }
    }

    this.setState({ isSubmitting: true });

    try {
      // Grade answers
      const gradedAnswers: IQuizAnswer[] = questions.map(question => {
        const userAnswer = answers.get(question.Id);
        if (userAnswer === undefined) {
          return {
            questionId: question.Id,
            questionType: question.QuestionType,
            selectedAnswer: "",
            isCorrect: false,
            isPartiallyCorrect: false,
            pointsEarned: 0,
            maxPoints: question.Points
          };
        }
        // Build the userResponse object expected by gradeAnswer
        const userResponse = this.buildGradeResponse(question, userAnswer);
        return this.quizService.gradeAnswer(question, userResponse);
      });

      // Submit attempt
      const result = await this.quizService.submitQuizAttempt(attemptId, gradedAnswers);

      if (!result) {
        throw new Error("Failed to submit quiz");
      }

      // Stop timer
      this.clearTimer();

      this.setState({
        result,
        showResults: true,
        isSubmitting: false
      });

      if (this.props.onComplete) {
        this.props.onComplete(result);
      }
    } catch (error) {
      console.error("Failed to submit quiz:", error);
      this.setState({
        error: "Failed to submit quiz. Please try again.",
        isSubmitting: false
      });
    }
  }

  /** Convert the raw answer value into the shape expected by QuizService.gradeAnswer */
  private buildGradeResponse(
    question: IQuizQuestion,
    answer: AnswerValue
  ): {
    selectedAnswer?: string;
    selectedAnswers?: string[];
    matchingAnswers?: { left: string; right: string }[];
    orderingAnswers?: string[];
    hotspotCoordinates?: { x: number; y: number };
    essayText?: string;
    ratingValue?: number;
    fillInBlanks?: string[];
  } {
    switch (question.QuestionType) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse:
      case QuestionType.ShortAnswer:
      case QuestionType.ImageChoice:
        return { selectedAnswer: answer as string };

      case QuestionType.MultipleSelect:
        return { selectedAnswers: answer as string[] };

      case QuestionType.FillInBlank:
        return { fillInBlanks: answer as string[] };

      case QuestionType.Matching:
        return { matchingAnswers: answer as { left: string; right: string }[] };

      case QuestionType.Ordering:
        return { orderingAnswers: answer as string[] };

      case QuestionType.RatingScale:
        return { ratingValue: answer as number };

      case QuestionType.Essay:
        return { essayText: answer as string };

      case QuestionType.Hotspot:
        return { hotspotCoordinates: answer as { x: number; y: number } };

      default:
        return { selectedAnswer: String(answer) };
    }
  }

  // ============================================================================
  // RENDER — QUESTION TYPES
  // ============================================================================

  private renderQuestion(question: IQuizQuestion, index: number): JSX.Element {
    const { answers } = this.state;
    const userAnswer = answers.get(question.Id);

    return (
      <div className={styles.questionContainer}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between">
            <Text variant="large" className={styles.questionNumber}>
              Question {index + 1} of {this.state.questions.length}
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <div className={styles.badge}>
                <Text variant="small">{question.DifficultyLevel}</Text>
              </div>
              <div className={styles.pointsBadge}>
                <Text variant="small">{question.Points} pts</Text>
              </div>
            </Stack>
          </Stack>

          <Text variant="xLarge" className={styles.questionText}>
            {question.QuestionText}
          </Text>

          {question.QuestionImage && (
            <img
              src={question.QuestionImage}
              alt="Question image"
              style={{ maxWidth: '100%', borderRadius: 8, marginBottom: 8 }}
            />
          )}

          {this.renderAnswerOptions(question, userAnswer)}

          {question.Hint && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
              Hint: {question.Hint}
            </MessageBar>
          )}
        </Stack>
      </div>
    );
  }

  private renderAnswerOptions(question: IQuizQuestion, userAnswer: AnswerValue | undefined): JSX.Element {
    switch (question.QuestionType) {
      case QuestionType.MultipleChoice:
        return this.renderMultipleChoice(question, userAnswer as string | undefined);

      case QuestionType.TrueFalse:
        return this.renderTrueFalse(question, userAnswer as string | undefined);

      case QuestionType.MultipleSelect:
        return this.renderMultipleSelect(question, userAnswer as string[] | undefined);

      case QuestionType.ShortAnswer:
        return this.renderShortAnswer(question, userAnswer as string | undefined);

      case QuestionType.FillInBlank:
        return this.renderFillInBlank(question, userAnswer as string[] | undefined);

      case QuestionType.Matching:
        return this.renderMatching(question, userAnswer as { left: string; right: string }[] | undefined);

      case QuestionType.Ordering:
        return this.renderOrdering(question, userAnswer as string[] | undefined);

      case QuestionType.RatingScale:
        return this.renderRatingScale(question, userAnswer as number | undefined);

      case QuestionType.Essay:
        return this.renderEssay(question, userAnswer as string | undefined);

      case QuestionType.ImageChoice:
        return this.renderImageChoice(question, userAnswer as string | undefined);

      case QuestionType.Hotspot:
        return this.renderHotspot(question, userAnswer as { x: number; y: number } | undefined);

      default:
        return <Text>Unsupported question type: {question.QuestionType}</Text>;
    }
  }

  // --- Multiple Choice ---
  private renderMultipleChoice(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const options: IChoiceGroupOption[] = [
      { key: "A", text: `A. ${question.OptionA || ''}` },
      { key: "B", text: `B. ${question.OptionB || ''}` },
      { key: "C", text: `C. ${question.OptionC || ''}` },
      { key: "D", text: `D. ${question.OptionD || ''}` },
      { key: "E", text: `E. ${question.OptionE || ''}` },
      { key: "F", text: `F. ${question.OptionF || ''}` }
    ].filter(opt => opt.text.length > 3); // Filter out empty options

    return (
      <ChoiceGroup
        selectedKey={userAnswer}
        options={options}
        onChange={(_ev, option) => {
          if (option) {
            this.handleAnswerChange(question.Id, option.key);
          }
        }}
        className={styles.choiceGroup}
      />
    );
  }

  // --- True/False ---
  private renderTrueFalse(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const options: IChoiceGroupOption[] = [
      { key: "A", text: "True" },
      { key: "B", text: "False" }
    ];

    return (
      <ChoiceGroup
        selectedKey={userAnswer}
        options={options}
        onChange={(_ev, option) => {
          if (option) {
            this.handleAnswerChange(question.Id, option.key);
          }
        }}
        className={styles.choiceGroup}
      />
    );
  }

  // --- Multiple Select ---
  private renderMultipleSelect(question: IQuizQuestion, userAnswers: string[] | undefined): JSX.Element {
    const options = [
      { key: "A", text: question.OptionA },
      { key: "B", text: question.OptionB },
      { key: "C", text: question.OptionC },
      { key: "D", text: question.OptionD },
      { key: "E", text: question.OptionE },
      { key: "F", text: question.OptionF }
    ].filter(opt => opt.text);

    const selectedAnswers = userAnswers || [];

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Select all that apply
        </MessageBar>
        {options.map(option => (
          <div
            key={option.key}
            className={`${styles.optionItem} ${selectedAnswers.includes(option.key) ? styles.selected : ''}`}
          >
            <Checkbox
              label={`${option.key}. ${option.text}`}
              checked={selectedAnswers.includes(option.key)}
              onChange={(_ev, checked) => {
                let newAnswers = [...selectedAnswers];
                if (checked) {
                  newAnswers.push(option.key);
                } else {
                  newAnswers = newAnswers.filter(a => a !== option.key);
                }
                this.handleAnswerChange(question.Id, newAnswers);
              }}
            />
          </div>
        ))}
      </Stack>
    );
  }

  // --- Short Answer ---
  private renderShortAnswer(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    return (
      <TextField
        multiline
        rows={4}
        value={userAnswer || ""}
        onChange={(_ev, value) => this.handleAnswerChange(question.Id, value || "")}
        placeholder="Type your answer here..."
      />
    );
  }

  // --- Fill in the Blank ---
  private renderFillInBlank(question: IQuizQuestion, userAnswers: string[] | undefined): JSX.Element {
    let blanks: { position: number; acceptedAnswers: string[] }[] = [];
    try {
      blanks = question.BlankAnswers ? JSON.parse(question.BlankAnswers) : [];
    } catch {
      blanks = [];
    }

    const currentAnswers = userAnswers || blanks.map(() => "");

    // Render the question text with blanks highlighted
    // QuestionText may contain "____" or "[blank]" markers
    const parts = question.QuestionText.split(/(?:____|\[blank\])/gi);

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Fill in the blank{blanks.length > 1 ? 's' : ''} below
        </MessageBar>
        <div className={styles.fillInBlankContainer}>
          {parts.map((part, index) => (
            <span key={index}>
              <span>{part}</span>
              {index < blanks.length && (
                <TextField
                  underlined
                  placeholder={`Blank ${index + 1}`}
                  value={currentAnswers[index] || ""}
                  onChange={(_ev, value) => {
                    const newAnswers = [...currentAnswers];
                    newAnswers[index] = value || "";
                    this.handleAnswerChange(question.Id, newAnswers);
                  }}
                  styles={{ root: { display: 'inline-block', width: 200, margin: '0 4px' } }}
                />
              )}
            </span>
          ))}
        </div>
        {question.CaseSensitive && (
          <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic' } }}>
            Answers are case-sensitive
          </Text>
        )}
      </Stack>
    );
  }

  // --- Matching ---
  private renderMatching(question: IQuizQuestion, userAnswers: { left: string; right: string }[] | undefined): JSX.Element {
    let pairs: { left: string; right: string }[] = [];
    try {
      pairs = question.MatchingPairs ? JSON.parse(question.MatchingPairs) : [];
    } catch {
      pairs = [];
    }

    // Create shuffled right-side options
    const rightOptions: IDropdownOption[] = pairs.map(p => ({
      key: p.right,
      text: p.right
    }));

    // Build current user matches
    const currentMatches = userAnswers || pairs.map(p => ({ left: p.left, right: "" }));

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Match each item on the left with the correct item on the right
        </MessageBar>
        {pairs.map((pair, index) => {
          const currentMatch = currentMatches.find(m => m.left === pair.left);
          return (
            <div key={index} className={styles.matchingRow}>
              <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center">
                <div className={styles.matchingLeft}>
                  <Text variant="medium">{pair.left}</Text>
                </div>
                <Text variant="large" styles={{ root: { color: '#0d9488', fontWeight: 600 } }}>→</Text>
                <Dropdown
                  placeholder="Select a match..."
                  selectedKey={currentMatch?.right || undefined}
                  options={rightOptions}
                  onChange={(_ev, option) => {
                    if (option) {
                      const newMatches = currentMatches.map(m =>
                        m.left === pair.left ? { left: m.left, right: option.key as string } : m
                      );
                      // Ensure all pairs have entries
                      const updatedMatches = pairs.map(p => {
                        const existing = newMatches.find(m => m.left === p.left);
                        return existing || { left: p.left, right: "" };
                      });
                      this.handleAnswerChange(question.Id, updatedMatches);
                    }
                  }}
                  styles={{ root: { minWidth: 200 } }}
                />
              </Stack>
            </div>
          );
        })}
      </Stack>
    );
  }

  // --- Ordering ---
  private renderOrdering(question: IQuizQuestion, userOrder: string[] | undefined): JSX.Element {
    let items: { id: string; text: string; correctOrder: number }[] = [];
    try {
      items = question.OrderingItems ? JSON.parse(question.OrderingItems) : [];
    } catch {
      items = [];
    }

    // Initialize order: use user's current order, or default shuffled order
    const currentOrder = userOrder || items.map(i => i.id);

    const orderedItems = currentOrder.map(id => items.find(i => i.id === id)).filter(Boolean) as typeof items;

    const moveItem = (fromIndex: number, toIndex: number): void => {
      const newOrder = [...currentOrder];
      const [moved] = newOrder.splice(fromIndex, 1);
      newOrder.splice(toIndex, 0, moved);
      this.handleAnswerChange(question.Id, newOrder);
    };

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Arrange the items in the correct order using the arrow buttons
        </MessageBar>
        {orderedItems.map((item, index) => (
          <div key={item.id} className={styles.orderingItem}>
            <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
              <div className={styles.orderNumber}>
                <Text variant="medium" styles={{ root: { fontWeight: 700, color: 'white' } }}>
                  {index + 1}
                </Text>
              </div>
              <Text variant="medium" styles={{ root: { flex: 1 } }}>{item.text}</Text>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <IconButton
                  iconProps={{ iconName: "ChevronUp" }}
                  disabled={index === 0}
                  onClick={() => moveItem(index, index - 1)}
                  title="Move up"
                  styles={{ root: { height: 32, width: 32 } }}
                />
                <IconButton
                  iconProps={{ iconName: "ChevronDown" }}
                  disabled={index === orderedItems.length - 1}
                  onClick={() => moveItem(index, index + 1)}
                  title="Move down"
                  styles={{ root: { height: 32, width: 32 } }}
                />
              </Stack>
            </Stack>
          </div>
        ))}
      </Stack>
    );
  }

  // --- Rating Scale ---
  private renderRatingScale(question: IQuizQuestion, userRating: number | undefined): JSX.Element {
    const min = question.ScaleMin ?? 1;
    const max = question.ScaleMax ?? 5;

    let labels: { min?: string; max?: string; mid?: string } = {};
    try {
      labels = question.ScaleLabels ? JSON.parse(question.ScaleLabels) : {};
    } catch {
      labels = {};
    }

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{labels.min || `${min}`}</Text>
          {labels.mid && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{labels.mid}</Text>}
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{labels.max || `${max}`}</Text>
        </Stack>
        <Slider
          min={min}
          max={max}
          step={1}
          value={userRating ?? min}
          showValue
          onChange={(value: number) => this.handleAnswerChange(question.Id, value)}
          styles={{ root: { marginTop: 8 } }}
        />
        <div className={styles.ratingDisplay}>
          <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: '#0d9488' } }}>
            {userRating !== undefined ? userRating : '-'}
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Selected value
          </Text>
        </div>
      </Stack>
    );
  }

  // --- Essay ---
  private renderEssay(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const wordCount = (userAnswer || "").trim().split(/\s+/).filter(w => w.length > 0).length;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          This question requires a written response and will be manually reviewed.
        </MessageBar>
        <TextField
          multiline
          rows={10}
          value={userAnswer || ""}
          onChange={(_ev, value) => this.handleAnswerChange(question.Id, value || "")}
          placeholder="Write your essay response here..."
          resizable
        />
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Word count: {wordCount}
          </Text>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {question.MinWordCount !== undefined && question.MinWordCount > 0 && (
              <Text variant="small" styles={{
                root: { color: wordCount < question.MinWordCount ? '#d13438' : '#107c10' }
              }}>
                Min: {question.MinWordCount} words
              </Text>
            )}
            {question.MaxWordCount !== undefined && question.MaxWordCount > 0 && (
              <Text variant="small" styles={{
                root: { color: wordCount > question.MaxWordCount ? '#d13438' : '#605e5c' }
              }}>
                Max: {question.MaxWordCount} words
              </Text>
            )}
          </Stack>
        </Stack>
      </Stack>
    );
  }

  // --- Image Choice ---
  private renderImageChoice(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const options = [
      { key: "A", text: question.OptionA, image: question.OptionAImage },
      { key: "B", text: question.OptionB, image: question.OptionBImage },
      { key: "C", text: question.OptionC, image: question.OptionCImage },
      { key: "D", text: question.OptionD, image: question.OptionDImage }
    ].filter(opt => opt.text || opt.image);

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Select the correct image
        </MessageBar>
        <div className={styles.imageChoiceGrid}>
          {options.map(option => (
            <div
              key={option.key}
              className={`${styles.imageChoiceCard} ${userAnswer === option.key ? styles.selected : ''}`}
              onClick={() => this.handleAnswerChange(question.Id, option.key)}
            >
              {option.image && (
                <img
                  src={option.image}
                  alt={`Option ${option.key}`}
                  className={styles.imageChoiceImg}
                />
              )}
              <div className={styles.imageChoiceLabel}>
                <Text variant="medium">{option.key}. {option.text || ''}</Text>
              </div>
              {userAnswer === option.key && (
                <div className={styles.imageChoiceCheck}>
                  <IconButton iconProps={{ iconName: "CheckMark" }} />
                </div>
              )}
            </div>
          ))}
        </div>
      </Stack>
    );
  }

  // --- Hotspot ---
  private renderHotspot(question: IQuizQuestion, userAnswer: { x: number; y: number } | undefined): JSX.Element {
    let hotspotData: { imageUrl: string; regions: { x: number; y: number; width: number; height: number; isCorrect: boolean }[] } | null = null;
    try {
      hotspotData = question.HotspotData ? JSON.parse(question.HotspotData) : null;
    } catch {
      hotspotData = null;
    }

    if (!hotspotData || !hotspotData.imageUrl) {
      return <Text>Hotspot question data is not configured properly.</Text>;
    }

    const handleImageClick = (e: React.MouseEvent<HTMLDivElement>): void => {
      const rect = e.currentTarget.getBoundingClientRect();
      const x = Math.round(((e.clientX - rect.left) / rect.width) * 100);
      const y = Math.round(((e.clientY - rect.top) / rect.height) * 100);
      this.handleAnswerChange(question.Id, { x, y });
    };

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Click on the correct area of the image
        </MessageBar>
        <div
          className={styles.hotspotContainer}
          onClick={handleImageClick}
          style={{ position: 'relative', cursor: 'crosshair' }}
        >
          <img
            src={hotspotData.imageUrl}
            alt="Hotspot question"
            style={{ width: '100%', display: 'block', borderRadius: 8 }}
          />
          {userAnswer && (
            <div
              className={styles.hotspotMarker}
              style={{
                position: 'absolute',
                left: `${userAnswer.x}%`,
                top: `${userAnswer.y}%`,
                transform: 'translate(-50%, -50%)'
              }}
            />
          )}
        </div>
        {userAnswer && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Click position: ({userAnswer.x}%, {userAnswer.y}%) — Click again to change
          </Text>
        )}
      </Stack>
    );
  }

  // ============================================================================
  // RENDER — CHROME (Timer, Nav, Results)
  // ============================================================================

  private renderTimer(): JSX.Element {
    const { timeRemaining } = this.state;
    const minutes = Math.floor(timeRemaining / 60);
    const seconds = timeRemaining % 60;
    const isLowTime = timeRemaining < 60;

    return (
      <div className={`${styles.timer} ${isLowTime ? styles.timerWarning : ''}`}>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <IconButton iconProps={{ iconName: "Clock" }} />
          <Text variant="large">
            {minutes}:{seconds < 10 ? '0' + seconds : seconds}
          </Text>
        </Stack>
      </div>
    );
  }

  private renderQuestionNav(): JSX.Element {
    const { questions, currentQuestionIndex, answers } = this.state;

    return (
      <div className={styles.questionNav}>
        <Text variant="medium" block className={styles.navTitle}>
          Question Navigation
        </Text>
        <div className={styles.questionGrid}>
          {questions.map((q, index) => {
            const isAnswered = answers.has(q.Id);
            const isCurrent = index === currentQuestionIndex;

            return (
              <button
                key={q.Id}
                className={`${styles.questionNavButton} ${isCurrent ? styles.current : ''} ${isAnswered ? styles.answered : ''}`}
                onClick={() => this.handleJumpToQuestion(index)}
              >
                {index + 1}
              </button>
            );
          })}
        </div>
        <Stack className={styles.navLegend} tokens={{ childrenGap: 8 }}>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <div className={`${styles.legendBox} ${styles.current}`} />
            <Text variant="small">Current</Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <div className={`${styles.legendBox} ${styles.answered}`} />
            <Text variant="small">Answered</Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <div className={styles.legendBox} />
            <Text variant="small">Not Answered</Text>
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderResults(): JSX.Element {
    const { result, quiz, questions } = this.state;

    if (!result || !quiz) return <></>;

    // Calculate remaining attempts correctly using attemptNumber, not attemptId
    const attemptsUsed = result.attemptId; // This is actually the attempt count for this user
    const remainingAttempts = Math.max(0, quiz.MaxAttempts - attemptsUsed);

    return (
      <div className={styles.resultsContainer}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Score Card */}
          <div className={`${styles.scoreCard} ${result.passed ? styles.passed : styles.failed}`}>
            <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
              <IconButton
                iconProps={{ iconName: result.passed ? "CompletedSolid" : "StatusCircleErrorX" }}
                className={styles.resultIcon}
              />
              <Text variant="xxLarge" className={styles.scoreText}>
                {result.percentage}%
              </Text>
              <Text variant="large">
                {result.passed ? "Passed!" : "Not Passed"}
              </Text>
              <Text variant="medium" className={styles.scoreDetails}>
                Score: {result.score} / {result.maxScore}
              </Text>
              <Text variant="small">
                Time: {result.timeSpent} minutes
              </Text>
              {result.passed && (
                <div className={styles.pointsEarned}>
                  <Text variant="medium">+{result.pointsEarned} points earned!</Text>
                </div>
              )}
              {result.requiresManualReview && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  {result.pendingQuestions} question(s) require manual review. Your final score may change.
                </MessageBar>
              )}
            </Stack>
          </div>

          {/* Pass/Fail Message */}
          <MessageBar messageBarType={result.passed ? MessageBarType.success : MessageBarType.error}>
            {result.passed
              ? `Congratulations! You passed with ${result.percentage}%. Passing score is ${quiz.PassingScore}%.`
              : `You need ${quiz.PassingScore}% to pass. You scored ${result.percentage}%. You have ${remainingAttempts} attempt(s) remaining.`}
          </MessageBar>

          {/* Answer Review */}
          {quiz.ShowCorrectAnswers && (
            <div className={styles.answerReview}>
              <Text variant="xLarge" block className={styles.reviewTitle}>
                Answer Review
              </Text>
              {questions.map((question, index) => {
                const answer = result.answers.find(a => a.questionId === question.Id);
                if (!answer) return null;

                return (
                  <div key={question.Id} className={styles.reviewItem}>
                    <Stack tokens={{ childrenGap: 12 }}>
                      <Stack horizontal horizontalAlign="space-between">
                        <Text variant="medium">
                          Question {index + 1}
                        </Text>
                        <div className={answer.isCorrect ? styles.correctBadge : (answer.isPartiallyCorrect ? styles.partialBadge : styles.incorrectBadge)}>
                          <Text variant="small">
                            {answer.isCorrect ? "Correct" : (answer.isPartiallyCorrect ? "Partial" : "Incorrect")}
                          </Text>
                        </div>
                      </Stack>

                      <Text>{question.QuestionText}</Text>

                      <Stack tokens={{ childrenGap: 8 }}>
                        <Text variant="small" className={styles.labelText}>
                          Your answer: <strong>{this.formatUserAnswer(answer)}</strong>
                        </Text>
                        {!answer.isCorrect && (
                          <Text variant="small" className={styles.labelText}>
                            Correct answer: <strong>{question.CorrectAnswer}</strong>
                          </Text>
                        )}
                        <Text variant="small" className={styles.labelText}>
                          Points: {answer.pointsEarned} / {answer.maxPoints}
                        </Text>
                      </Stack>

                      {answer.feedback && (
                        <MessageBar messageBarType={MessageBarType.info}>
                          {answer.feedback}
                        </MessageBar>
                      )}

                      {question.Explanation && !answer.feedback && (
                        <MessageBar messageBarType={MessageBarType.info}>
                          {question.Explanation}
                        </MessageBar>
                      )}
                    </Stack>
                  </div>
                );
              })}
            </div>
          )}

          {/* Actions */}
          <Stack horizontal tokens={{ childrenGap: 16 }} horizontalAlign="center">
            <PrimaryButton text="Close" onClick={this.handleCancel} />
            {!result.passed && remainingAttempts > 0 && (
              <DefaultButton
                text="Retake Quiz"
                onClick={() => this.retakeQuiz()}
              />
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  /** Format a graded answer for display in the review */
  private formatUserAnswer(answer: IQuizAnswer): string {
    if (answer.essayText) return answer.essayText.substring(0, 100) + (answer.essayText.length > 100 ? '...' : '');
    if (answer.selectedAnswers && answer.selectedAnswers.length > 0) return answer.selectedAnswers.join(', ');
    if (answer.matchingAnswers) return answer.matchingAnswers.map(m => `${m.left} → ${m.right}`).join(', ');
    if (answer.orderingAnswers) return answer.orderingAnswers.join(' → ');
    if (answer.ratingValue !== undefined) return String(answer.ratingValue);
    if (answer.hotspotCoordinates) return `(${answer.hotspotCoordinates.x}%, ${answer.hotspotCoordinates.y}%)`;
    return answer.selectedAnswer || "(no answer)";
  }

  // ============================================================================
  // MAIN RENDER
  // ============================================================================

  public render(): React.ReactElement<IQuizTakerProps> {
    const {
      quiz,
      questions,
      currentQuestionIndex,
      error,
      isSubmitting,
      isLoading,
      showResults,
      answers
    } = this.state;

    if (error && !showResults) {
      return (
        <div className={styles.quizTaker}>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
          <DefaultButton text="Close" onClick={this.handleCancel} style={{ marginTop: 16 }} />
        </div>
      );
    }

    if (isLoading || !quiz || questions.length === 0) {
      return (
        <div className={styles.quizTaker}>
          <Stack horizontalAlign="center" tokens={{ childrenGap: 16, padding: 48 }}>
            <Spinner size={SpinnerSize.large} label="Loading quiz..." />
          </Stack>
        </div>
      );
    }

    if (showResults) {
      return (
        <div className={styles.quizTaker}>
          {this.renderResults()}
        </div>
      );
    }

    const currentQuestion = questions[currentQuestionIndex];
    const progress = ((currentQuestionIndex + 1) / questions.length) * 100;
    const answeredCount = answers.size;

    return (
      <div className={styles.quizTaker}>
        <div className={styles.header}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="xxLarge">{quiz.Title}</Text>
              <Text variant="small">{quiz.QuizDescription}</Text>
            </Stack>
            {this.renderTimer()}
          </Stack>
        </div>

        <ProgressIndicator
          label={`Progress: ${answeredCount} / ${questions.length} answered`}
          percentComplete={progress / 100}
        />

        <div className={styles.content}>
          <div className={styles.mainContent}>
            {this.renderQuestion(currentQuestion, currentQuestionIndex)}

            <Stack horizontal tokens={{ childrenGap: 16 }} horizontalAlign="space-between" className={styles.navigation}>
              <DefaultButton
                text="Previous"
                iconProps={{ iconName: "ChevronLeft" }}
                disabled={currentQuestionIndex === 0}
                onClick={this.handlePrevious}
              />

              <Stack horizontal tokens={{ childrenGap: 16 }}>
                {currentQuestionIndex === questions.length - 1 ? (
                  <PrimaryButton
                    text="Submit Quiz"
                    iconProps={{ iconName: "CheckMark" }}
                    onClick={() => this.handleSubmit()}
                    disabled={isSubmitting}
                  />
                ) : (
                  <PrimaryButton
                    text="Next"
                    iconProps={{ iconName: "ChevronRight" }}
                    onClick={this.handleNext}
                  />
                )}

                <DefaultButton
                  text="Cancel"
                  onClick={this.handleCancel}
                  disabled={isSubmitting}
                />
              </Stack>
            </Stack>
          </div>

          <div className={styles.sidebar}>
            {this.renderQuestionNav()}
          </div>
        </div>
      </div>
    );
  }
}
