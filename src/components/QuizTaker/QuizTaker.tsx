// @ts-nocheck
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
  Panel,
  PanelType,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  TextField,
  Label,
  IStackTokens,
  mergeStyleSets
} from '@fluentui/react';
import { QuizService, IQuiz, IQuizQuestion, IQuizAnswer, IQuizResult } from '../../services/QuizService';
import { SPFI } from '@pnp/sp';

export interface IQuizTakerProps {
  sp: SPFI;
  quizId: number;
  policyId: number;
  userId: any;
  onComplete?: (result: IQuizResult) => void;
  onCancel?: () => void;
}

export interface IQuizTakerState {
  quiz: IQuiz | null;
  questions: IQuizQuestion[];
  currentQuestionIndex: number;
  answers: Map<number, string | string[]>;
  timeRemaining: number;
  attemptId: number | null;
  isSubmitting: boolean;
  result: IQuizResult | null;
  error: string | null;
  showResults: boolean;
  startTime: Date;
}

export class QuizTaker extends React.Component<IQuizTakerProps, IQuizTakerState> {
  private quizService: QuizService;
  private timerInterval: any;

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
    if (this.timerInterval) {
      clearInterval(this.timerInterval);
    }
  }

  private async loadQuiz(): Promise<void> {
    try {
      const quiz = await this.quizService.getQuizById(this.props.quizId);
      if (!quiz) {
        this.setState({ error: "Quiz not found" });
        return;
      }

      // Check eligibility
      const eligibility = await this.quizService.canUserTakeQuiz(this.props.quizId, this.props.userId.Id);
      if (!eligibility.canTake) {
        this.setState({ error: eligibility.reason || "Cannot take quiz" });
        return;
      }

      // Get questions
      const questions = await this.quizService.getQuizQuestions(this.props.quizId, { randomize: quiz.RandomizeQuestions });

      // Start attempt
      const userId = this.props.userId?.Id || this.props.userId;
      const userName = this.props.userId?.Title || this.props.userId?.DisplayName || 'Unknown User';
      const userEmail = this.props.userId?.EMail || this.props.userId?.Email || '';
      const attempt = await this.quizService.startQuizAttempt(
        this.props.quizId,
        this.props.policyId,
        userId,
        userName,
        userEmail
      );

      if (!attempt) {
        this.setState({ error: "Failed to start quiz attempt" });
        return;
      }

      this.setState({
        quiz,
        questions,
        attemptId: attempt.Id,
        timeRemaining: quiz.TimeLimit * 60, // Convert to seconds
        startTime: new Date()
      });

      // Start timer
      this.startTimer();
    } catch (error) {
      console.error("Failed to load quiz:", error);
      this.setState({ error: "Failed to load quiz" });
    }
  }

  private startTimer(): void {
    this.timerInterval = setInterval(() => {
      this.setState(prev => {
        const newTimeRemaining = prev.timeRemaining - 1;

        if (newTimeRemaining <= 0) {
          clearInterval(this.timerInterval);
          this.handleTimeExpired();
          return { timeRemaining: 0 };
        }

        return { timeRemaining: newTimeRemaining };
      });
    }, 1000);
  }

  private handleTimeExpired(): void {
    this.setState({ error: "Time expired! Quiz will be submitted automatically." });
    setTimeout(() => {
      this.handleSubmit();
    }, 2000);
  }

  private handleAnswerChange = (questionId: number, answer: string | string[]): void => {
    this.setState(prev => {
      const newAnswers = new Map(prev.answers);
      newAnswers.set(questionId, answer);
      return { answers: newAnswers };
    });
  };

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
        if (!userAnswer) {
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
        // Convert the simple answer to the expected format
        const userResponse = typeof userAnswer === 'string'
          ? { selectedAnswer: userAnswer }
          : { selectedAnswers: userAnswer };
        return this.quizService.gradeAnswer(question, userResponse);
      });

      // Submit attempt
      const result = await this.quizService.submitQuizAttempt(attemptId, gradedAnswers);

      if (!result) {
        throw new Error("Failed to submit quiz");
      }

      // Stop timer
      if (this.timerInterval) {
        clearInterval(this.timerInterval);
      }

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

          {this.renderAnswerOptions(question, userAnswer)}
        </Stack>
      </div>
    );
  }

  private renderAnswerOptions(question: IQuizQuestion, userAnswer: string | string[] | undefined): JSX.Element {
    switch (question.QuestionType) {
      case "Multiple Choice":
        return this.renderMultipleChoice(question, userAnswer as string);

      case "True/False":
        return this.renderTrueFalse(question, userAnswer as string);

      case "Multiple Select":
        return this.renderMultipleSelect(question, userAnswer as string[]);

      case "Short Answer":
        return this.renderShortAnswer(question, userAnswer as string);

      default:
        return <Text>Unknown question type</Text>;
    }
  }

  private renderMultipleChoice(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const options: IChoiceGroupOption[] = [
      { key: "A", text: `A. ${question.OptionA}` },
      { key: "B", text: `B. ${question.OptionB}` },
      { key: "C", text: `C. ${question.OptionC}` },
      { key: "D", text: `D. ${question.OptionD}` }
    ].filter(opt => opt.text && opt.text.length > 3); // Filter out empty options (just "A. " etc.)

    return (
      <ChoiceGroup
        selectedKey={userAnswer}
        options={options}
        onChange={(_, option) => {
          if (option) {
            this.handleAnswerChange(question.Id, option.key);
          }
        }}
        className={styles.choiceGroup}
      />
    );
  }

  private renderTrueFalse(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    const options: IChoiceGroupOption[] = [
      { key: "A", text: "True" },
      { key: "B", text: "False" }
    ];

    return (
      <ChoiceGroup
        selectedKey={userAnswer}
        options={options}
        onChange={(_, option) => {
          if (option) {
            this.handleAnswerChange(question.Id, option.key);
          }
        }}
        className={styles.choiceGroup}
      />
    );
  }

  private renderMultipleSelect(question: IQuizQuestion, userAnswers: string[] | undefined): JSX.Element {
    const options = [
      { key: "A", text: question.OptionA },
      { key: "B", text: question.OptionB },
      { key: "C", text: question.OptionC },
      { key: "D", text: question.OptionD }
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
              onChange={(e, checked) => {
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

  private renderShortAnswer(question: IQuizQuestion, userAnswer: string | undefined): JSX.Element {
    return (
      <TextField
        multiline
        rows={4}
        value={userAnswer || ""}
        onChange={(e, value) => this.handleAnswerChange(question.Id, value || "")}
        placeholder="Type your answer here..."
      />
    );
  }

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
            {minutes}:{seconds.toString().padStart(2, '0')}
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
            <div className={`${styles.legendBox} ${styles.current}`}></div>
            <Text variant="small">Current</Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <div className={`${styles.legendBox} ${styles.answered}`}></div>
            <Text variant="small">Answered</Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <div className={`${styles.legendBox}`}></div>
            <Text variant="small">Not Answered</Text>
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderResults(): JSX.Element {
    const { result, quiz, questions } = this.state;

    if (!result || !quiz) return <></>;

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
                  <Text variant="medium">🏆 +{result.pointsEarned} points earned!</Text>
                </div>
              )}
            </Stack>
          </div>

          {/* Pass/Fail Message */}
          <MessageBar messageBarType={result.passed ? MessageBarType.success : MessageBarType.error}>
            {result.passed
              ? `Congratulations! You passed with ${result.percentage}%. Passing score is ${quiz.PassingScore}%.`
              : `You need ${quiz.PassingScore}% to pass. You scored ${result.percentage}%. You have ${quiz.MaxAttempts - result.attemptId} attempt(s) remaining.`}
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
                        <div className={answer.isCorrect ? styles.correctBadge : styles.incorrectBadge}>
                          <Text variant="small">
                            {answer.isCorrect ? "✓ Correct" : "✗ Incorrect"}
                          </Text>
                        </div>
                      </Stack>

                      <Text>{question.QuestionText}</Text>

                      <Stack tokens={{ childrenGap: 8 }}>
                        <Text variant="small" className={styles.labelText}>
                          Your answer: <strong>{answer.selectedAnswer}</strong>
                        </Text>
                        {!answer.isCorrect && (
                          <Text variant="small" className={styles.labelText}>
                            Correct answer: <strong>{question.CorrectAnswer}</strong>
                          </Text>
                        )}
                      </Stack>

                      {question.Explanation && (
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
            <PrimaryButton text="Close" onClick={this.props.onCancel} />
            {!result.passed && (
              <DefaultButton text="Retake Quiz" onClick={() => window.location.reload()} />
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  public render(): React.ReactElement<IQuizTakerProps> {
    const {
      quiz,
      questions,
      currentQuestionIndex,
      error,
      isSubmitting,
      showResults,
      answers
    } = this.state;

    if (error) {
      return (
        <div className={styles.quizTaker}>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
          <DefaultButton text="Close" onClick={this.props.onCancel} style={{ marginTop: 16 }} />
        </div>
      );
    }

    if (!quiz || questions.length === 0) {
      return <div className={styles.quizTaker}>Loading quiz...</div>;
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
    const answeredCount = Array.from(answers.keys()).length;

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
                  onClick={this.props.onCancel}
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
