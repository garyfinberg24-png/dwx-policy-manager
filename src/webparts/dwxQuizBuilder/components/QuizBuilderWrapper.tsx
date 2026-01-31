// @ts-nocheck
import * as React from 'react';
import { IQuizBuilderWrapperProps } from './IQuizBuilderWrapperProps';
import { QuizBuilder } from '../../../components/QuizBuilder';
import { DwxAppLayout } from '../../../components/JmlAppLayout';
import { QuizService, IQuiz, QuizStatus } from '../../../services/QuizService';
import styles from './QuizBuilderWrapper.module.scss';

// Status badge colors
const statusColors: Record<string, { bg: string; color: string }> = {
  Draft: { bg: '#fff4ce', color: '#8a6d00' },
  Published: { bg: '#dff6dd', color: '#107c10' },
  Scheduled: { bg: '#cce4f6', color: '#0078d4' },
  Archived: { bg: '#f3f2f1', color: '#605e5c' }
};

const difficultyColors: Record<string, { bg: string; color: string }> = {
  Easy: { bg: '#dff6dd', color: '#107c10' },
  Medium: { bg: '#fff4ce', color: '#8a6d00' },
  Hard: { bg: '#fed9cc', color: '#d83b01' },
  Expert: { bg: '#fde7e9', color: '#a80000' }
};

export const QuizBuilderWrapper: React.FC<IQuizBuilderWrapperProps> = (props) => {
  const { sp, context, title } = props;

  // Get quizId and policyId from URL params if present
  const urlParams = new URLSearchParams(window.location.search);
  const quizId = urlParams.get('quizId') ? parseInt(urlParams.get('quizId')!, 10) : undefined;
  const policyId = urlParams.get('policyId') ? parseInt(urlParams.get('policyId')!, 10) : undefined;

  // Quiz list state
  const [quizzes, setQuizzes] = React.useState<IQuiz[]>([]);
  const [listLoading, setListLoading] = React.useState(true);
  const [listError, setListError] = React.useState<string | null>(null);
  const [filterStatus, setFilterStatus] = React.useState<string>('all');
  const [searchText, setSearchText] = React.useState('');

  const quizService = React.useMemo(() => new QuizService(sp), [sp]);

  // Load quizzes when in list mode (no quizId)
  React.useEffect(() => {
    if (!quizId) {
      loadQuizzes();
    }
  }, [quizId]);

  const loadQuizzes = async (): Promise<void> => {
    setListLoading(true);
    setListError(null);
    try {
      const allQuizzes = await quizService.getAllQuizzes({ includeArchived: true });
      setQuizzes(allQuizzes);
    } catch (err) {
      console.error('[QuizList] Failed to load quizzes:', err);
      setListError('Failed to load quizzes. Ensure the PM_PolicyQuizzes list exists.');
    } finally {
      setListLoading(false);
    }
  };

  const handleSave = (quiz: any): void => {
    console.log('Quiz saved:', quiz);
  };

  const handleCancel = (): void => {
    // Navigate back to quiz list
    window.location.href = '/sites/PolicyManager/SitePages/QuizBuilder.aspx';
  };

  const handleEditQuiz = (id: number): void => {
    window.location.href = `/sites/PolicyManager/SitePages/QuizBuilder.aspx?quizId=${id}`;
  };

  const handleDuplicateQuiz = async (quiz: IQuiz): Promise<void> => {
    try {
      const newQuiz = await quizService.createQuiz({
        Title: `${quiz.Title} (Copy)`,
        PolicyId: quiz.PolicyId,
        PolicyTitle: quiz.PolicyTitle,
        QuizDescription: quiz.QuizDescription,
        PassingScore: quiz.PassingScore,
        TimeLimit: quiz.TimeLimit,
        MaxAttempts: quiz.MaxAttempts,
        QuizCategory: quiz.QuizCategory,
        DifficultyLevel: quiz.DifficultyLevel,
        RandomizeQuestions: quiz.RandomizeQuestions,
        ShowCorrectAnswers: quiz.ShowCorrectAnswers,
        AllowPartialCredit: quiz.AllowPartialCredit,
        GenerateCertificate: quiz.GenerateCertificate,
        Status: QuizStatus.Draft
      });
      if (newQuiz) {
        // Navigate to edit the new quiz
        window.location.href = `/sites/PolicyManager/SitePages/QuizBuilder.aspx?quizId=${newQuiz.Id}`;
      }
    } catch (err) {
      console.error('[QuizList] Failed to duplicate quiz:', err);
    }
  };

  // Filter and search quizzes
  const filteredQuizzes = React.useMemo(() => {
    let list = quizzes;
    if (filterStatus !== 'all') {
      list = list.filter(q => q.Status === filterStatus);
    }
    if (searchText.trim()) {
      const lower = searchText.toLowerCase();
      list = list.filter(q =>
        q.Title?.toLowerCase().includes(lower) ||
        q.PolicyTitle?.toLowerCase().includes(lower) ||
        q.QuizCategory?.toLowerCase().includes(lower)
      );
    }
    return list;
  }, [quizzes, filterStatus, searchText]);

  // Breadcrumbs
  const breadcrumbs = quizId
    ? [
        { text: 'Policy Manager', href: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
        { text: 'Quiz Builder', href: '/sites/PolicyManager/SitePages/QuizBuilder.aspx' },
        { text: `Quiz #${quizId}` }
      ]
    : [
        { text: 'Policy Manager', href: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
        { text: 'Quiz Builder' }
      ];

  // ====================================================================
  // QUIZ LIST VIEW (no quizId — show all quizzes)
  // ====================================================================
  if (!quizId) {
    return (
      <DwxAppLayout
        title={title}
        context={context}
        showBreadcrumb={true}
        fullWidth={true}
        breadcrumbs={breadcrumbs}
        activeNavKey="quiz"
      >
        <div className={styles.quizBuilderWrapper}>
          <div style={{ maxWidth: '1400px', margin: '0 auto', padding: '20px' }}>
            {/* Header */}
            <div style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              marginBottom: '24px',
              flexWrap: 'wrap',
              gap: '12px'
            }}>
              <div>
                <h1 style={{ fontSize: '28px', fontWeight: 600, color: '#323130', margin: 0 }}>Quiz Builder</h1>
                <p style={{ fontSize: '14px', color: '#605e5c', margin: '4px 0 0' }}>
                  Create, manage and edit policy compliance quizzes
                </p>
              </div>
              <button
                onClick={() => window.location.href = '/sites/PolicyManager/SitePages/QuizBuilder.aspx?quizId=new'}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 20px',
                  backgroundColor: '#0d9488',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '6px',
                  fontSize: '14px',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                <svg viewBox="0 0 24 24" fill="none" style={{ width: 18, height: 18 }}>
                  <path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"/>
                </svg>
                Create New Quiz
              </button>
            </div>

            {/* Filter & Search Bar */}
            <div style={{
              display: 'flex',
              gap: '12px',
              marginBottom: '20px',
              flexWrap: 'wrap',
              alignItems: 'center'
            }}>
              <div style={{
                display: 'flex',
                gap: '4px',
                backgroundColor: '#f3f2f1',
                borderRadius: '6px',
                padding: '3px'
              }}>
                {['all', 'Draft', 'Published', 'Scheduled', 'Archived'].map(status => (
                  <button
                    key={status}
                    onClick={() => setFilterStatus(status)}
                    style={{
                      padding: '6px 14px',
                      fontSize: '13px',
                      fontWeight: filterStatus === status ? 600 : 400,
                      border: 'none',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      backgroundColor: filterStatus === status ? '#fff' : 'transparent',
                      color: filterStatus === status ? '#0d9488' : '#605e5c',
                      boxShadow: filterStatus === status ? '0 1px 3px rgba(0,0,0,0.12)' : 'none',
                      transition: 'all 0.15s'
                    }}
                  >
                    {status === 'all' ? 'All' : status}
                  </button>
                ))}
              </div>
              <div style={{ flex: 1, minWidth: '200px', position: 'relative' }}>
                <svg viewBox="0 0 24 24" fill="none" style={{
                  width: 16, height: 16, position: 'absolute', left: '12px', top: '50%',
                  transform: 'translateY(-50%)', color: '#a19f9d'
                }}>
                  <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                </svg>
                <input
                  type="text"
                  placeholder="Search quizzes..."
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                  style={{
                    width: '100%',
                    padding: '8px 12px 8px 36px',
                    border: '1px solid #d2d0ce',
                    borderRadius: '6px',
                    fontSize: '14px',
                    outline: 'none',
                    boxSizing: 'border-box'
                  }}
                />
              </div>
              <span style={{ fontSize: '13px', color: '#605e5c' }}>
                {filteredQuizzes.length} quiz{filteredQuizzes.length !== 1 ? 'zes' : ''}
              </span>
            </div>

            {/* Loading / Error */}
            {listLoading && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: '#605e5c' }}>
                <div style={{ fontSize: '14px' }}>Loading quizzes...</div>
              </div>
            )}

            {listError && (
              <div style={{
                padding: '16px',
                backgroundColor: '#fde7e9',
                border: '1px solid #a80000',
                borderRadius: '6px',
                color: '#a80000',
                fontSize: '14px',
                marginBottom: '16px'
              }}>
                {listError}
              </div>
            )}

            {/* Empty State */}
            {!listLoading && !listError && filteredQuizzes.length === 0 && (
              <div style={{
                textAlign: 'center',
                padding: '60px 20px',
                backgroundColor: '#fff',
                borderRadius: '8px',
                border: '1px solid #edebe9'
              }}>
                <svg viewBox="0 0 24 24" fill="none" style={{ width: 48, height: 48, color: '#c8c6c4', marginBottom: '12px' }}>
                  <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
                <h3 style={{ color: '#323130', margin: '0 0 8px' }}>
                  {searchText || filterStatus !== 'all' ? 'No quizzes match your filters' : 'No quizzes yet'}
                </h3>
                <p style={{ color: '#605e5c', margin: '0 0 16px', fontSize: '14px' }}>
                  {searchText || filterStatus !== 'all'
                    ? 'Try adjusting your search or filter criteria.'
                    : 'Create your first policy compliance quiz to get started.'}
                </p>
                {!searchText && filterStatus === 'all' && (
                  <button
                    onClick={() => window.location.href = '/sites/PolicyManager/SitePages/QuizBuilder.aspx?quizId=new'}
                    style={{
                      padding: '10px 20px',
                      backgroundColor: '#0d9488',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '6px',
                      fontSize: '14px',
                      fontWeight: 600,
                      cursor: 'pointer'
                    }}
                  >
                    Create First Quiz
                  </button>
                )}
              </div>
            )}

            {/* Quiz Cards Grid */}
            {!listLoading && filteredQuizzes.length > 0 && (
              <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))',
                gap: '16px'
              }}>
                {filteredQuizzes.map(quiz => {
                  const sColor = statusColors[quiz.Status] || statusColors.Draft;
                  const dColor = difficultyColors[quiz.DifficultyLevel] || difficultyColors.Medium;
                  return (
                    <div
                      key={quiz.Id}
                      style={{
                        backgroundColor: '#fff',
                        borderRadius: '8px',
                        border: '1px solid #edebe9',
                        padding: '20px',
                        transition: 'box-shadow 0.2s, transform 0.2s',
                        cursor: 'pointer',
                        display: 'flex',
                        flexDirection: 'column'
                      }}
                      onClick={() => handleEditQuiz(quiz.Id)}
                      onMouseEnter={(e) => {
                        (e.currentTarget as HTMLElement).style.boxShadow = '0 4px 12px rgba(0,0,0,0.12)';
                        (e.currentTarget as HTMLElement).style.transform = 'translateY(-2px)';
                      }}
                      onMouseLeave={(e) => {
                        (e.currentTarget as HTMLElement).style.boxShadow = 'none';
                        (e.currentTarget as HTMLElement).style.transform = 'none';
                      }}
                    >
                      {/* Card Header */}
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
                        <div style={{ flex: 1 }}>
                          <h3 style={{ fontSize: '16px', fontWeight: 600, color: '#323130', margin: '0 0 4px' }}>
                            {quiz.Title}
                          </h3>
                          {quiz.PolicyTitle && (
                            <div style={{ fontSize: '12px', color: '#605e5c' }}>
                              Linked: {quiz.PolicyTitle}
                            </div>
                          )}
                        </div>
                        <span style={{
                          fontSize: '11px',
                          fontWeight: 500,
                          padding: '2px 8px',
                          borderRadius: '12px',
                          backgroundColor: sColor.bg,
                          color: sColor.color,
                          flexShrink: 0
                        }}>
                          {quiz.Status}
                        </span>
                      </div>

                      {/* Description */}
                      {quiz.QuizDescription && (
                        <p style={{
                          fontSize: '13px',
                          color: '#605e5c',
                          margin: '0 0 12px',
                          lineHeight: '1.4',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          display: '-webkit-box',
                          WebkitLineClamp: 2,
                          WebkitBoxOrient: 'vertical'
                        }}>
                          {quiz.QuizDescription}
                        </p>
                      )}

                      {/* Stats Row */}
                      <div style={{
                        display: 'flex',
                        gap: '16px',
                        fontSize: '13px',
                        color: '#605e5c',
                        marginBottom: '12px',
                        flexWrap: 'wrap'
                      }}>
                        <span title="Questions">
                          <strong>{quiz.QuestionCount || 0}</strong> questions
                        </span>
                        <span title="Passing Score">
                          Pass: <strong>{quiz.PassingScore}%</strong>
                        </span>
                        <span title="Time Limit">
                          <strong>{quiz.TimeLimit}</strong> min
                        </span>
                        <span title="Max Attempts">
                          <strong>{quiz.MaxAttempts}</strong> attempts
                        </span>
                      </div>

                      {/* Tags */}
                      <div style={{ display: 'flex', gap: '6px', marginBottom: '16px', flexWrap: 'wrap' }}>
                        {quiz.QuizCategory && (
                          <span style={{
                            fontSize: '11px',
                            padding: '2px 8px',
                            borderRadius: '12px',
                            backgroundColor: '#f3f2f1',
                            color: '#605e5c'
                          }}>
                            {quiz.QuizCategory}
                          </span>
                        )}
                        {quiz.DifficultyLevel && (
                          <span style={{
                            fontSize: '11px',
                            padding: '2px 8px',
                            borderRadius: '12px',
                            backgroundColor: dColor.bg,
                            color: dColor.color
                          }}>
                            {quiz.DifficultyLevel}
                          </span>
                        )}
                        {quiz.GenerateCertificate && (
                          <span style={{
                            fontSize: '11px',
                            padding: '2px 8px',
                            borderRadius: '12px',
                            backgroundColor: '#dff6dd',
                            color: '#107c10'
                          }}>
                            Certificate
                          </span>
                        )}
                      </div>

                      {/* Action Buttons */}
                      <div style={{ marginTop: 'auto', display: 'flex', gap: '8px', borderTop: '1px solid #f3f2f1', paddingTop: '12px' }}>
                        <button
                          onClick={(e) => { e.stopPropagation(); handleEditQuiz(quiz.Id); }}
                          style={{
                            flex: 1,
                            padding: '8px',
                            fontSize: '13px',
                            fontWeight: 500,
                            border: '1px solid #0d9488',
                            borderRadius: '4px',
                            backgroundColor: '#0d9488',
                            color: '#fff',
                            cursor: 'pointer'
                          }}
                        >
                          Edit
                        </button>
                        <button
                          onClick={(e) => { e.stopPropagation(); void handleDuplicateQuiz(quiz); }}
                          style={{
                            flex: 1,
                            padding: '8px',
                            fontSize: '13px',
                            fontWeight: 500,
                            border: '1px solid #d2d0ce',
                            borderRadius: '4px',
                            backgroundColor: '#fff',
                            color: '#323130',
                            cursor: 'pointer'
                          }}
                        >
                          Duplicate
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </DwxAppLayout>
    );
  }

  // ====================================================================
  // QUIZ EDITOR VIEW (quizId present)
  // ====================================================================
  return (
    <DwxAppLayout
      title={title}
      context={context}
      showBreadcrumb={true}
      fullWidth={true}
      breadcrumbs={breadcrumbs}
      activeNavKey="quiz"
    >
      <div className={styles.quizBuilderWrapper}>
        <QuizBuilder
          sp={sp}
          context={context}
          quizId={quizId === 'new' as any ? undefined : quizId}
          policyId={policyId}
          aiFunctionUrl={props.aiFunctionUrl}
          onSave={handleSave}
          onCancel={handleCancel}
        />
      </div>
    </DwxAppLayout>
  );
};

export default QuizBuilderWrapper;
