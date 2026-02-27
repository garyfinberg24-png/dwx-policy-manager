// @ts-nocheck
/**
 * QuizBuilderTab — Extracted from PolicyAuthorEnhanced.tsx
 * Displays quiz management interface with quick-create cards and quiz table.
 */
import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  IconButton,
  Icon,
} from '@fluentui/react';
import { PageSubheader } from '../../../../components/PageSubheader';
import { PrimaryButton } from '@fluentui/react';
import { IQuizBuilderTabProps } from './types';

export default class QuizBuilderTab extends React.Component<IQuizBuilderTabProps> {

  public render(): React.ReactElement<IQuizBuilderTabProps> {
    const { quizzes, quizzesLoading, styles, dialogManager, onCreateQuiz, onEditQuiz } = this.props;

    return (
      <>
        <PageSubheader
          iconName="Questionnaire"
          title="Quiz Builder"
          description="Create quizzes to verify policy understanding"
          actions={
            <PrimaryButton
              text="Create New Quiz"
              iconProps={{ iconName: 'Add' }}
              onClick={onCreateQuiz}
            />
          }
        />

        <div className={styles.editorContainer}>
          {quizzesLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading quizzes..." />
            </Stack>
          ) : (
            <>
              {/* Quick Create Section */}
              <div className={styles.quickCreateSection}>
                <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Quick Create Quiz</Text>
                <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                  <div className={styles.quickCreateCard} onClick={onCreateQuiz}>
                    <div className={styles.quickCreateIcon} style={{ background: '#e8f4fd' }}>
                      <Icon iconName="Questionnaire" style={{ fontSize: 24, color: '#0078d4' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>From Scratch</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Create a new quiz manually</Text>
                    </div>
                  </div>
                  <div className={styles.quickCreateCard} onClick={() => window.location.href = '/sites/PolicyManager/SitePages/QuizBuilder.aspx'}>
                    <div className={styles.quickCreateIcon} style={{ background: '#f3e8fd' }}>
                      <Icon iconName="Robot" style={{ fontSize: 24, color: '#8764b8' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>AI Generated</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Auto-generate from policy content</Text>
                    </div>
                  </div>
                  <div className={styles.quickCreateCard} onClick={() => void dialogManager.showAlert('Quiz template library coming soon. Use AI Generated or From Scratch.', { variant: 'info' })}>
                    <div className={styles.quickCreateIcon} style={{ background: '#dff6dd' }}>
                      <Icon iconName="DocumentSet" style={{ fontSize: 24, color: '#107c10' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>From Template</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Use an existing quiz template</Text>
                    </div>
                  </div>
                </Stack>
              </div>

              {/* Quiz Table */}
              <div style={{ marginTop: 24 }}>
                <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Existing Quizzes</Text>
                <div className={styles.quizTable}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f3f2f1', borderBottom: '2px solid #edebe9' }}>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Quiz Title</th>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Linked Policy</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Questions</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Pass Rate</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Status</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Completions</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Avg Score</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {quizzes.map((quiz, index) => (
                        <tr key={quiz.Id} style={{ borderBottom: '1px solid #edebe9', background: index % 2 === 0 ? '#ffffff' : '#faf9f8' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 500 }}>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                              <Icon iconName="Questionnaire" style={{ color: '#0078d4' }} />
                              <span>{quiz.Title}</span>
                            </Stack>
                          </td>
                          <td style={{ padding: '12px 16px', color: '#605e5c' }}>{quiz.LinkedPolicy}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.Questions}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.PassRate}%</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              display: 'inline-block',
                              padding: '4px 12px',
                              borderRadius: '12px',
                              fontSize: '11px',
                              fontWeight: 600,
                              textTransform: 'uppercase',
                              background: quiz.Status === 'Active' ? '#dff6dd' : quiz.Status === 'Draft' ? '#fff4ce' : '#f3f2f1',
                              color: quiz.Status === 'Active' ? '#107c10' : quiz.Status === 'Draft' ? '#8a6d3b' : '#605e5c'
                            }}>
                              {quiz.Status}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.Completions}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            {quiz.AvgScore > 0 ? (
                              <span style={{ color: quiz.AvgScore >= 80 ? '#107c10' : quiz.AvgScore >= 60 ? '#ca5010' : '#d13438' }}>
                                {quiz.AvgScore}%
                              </span>
                            ) : '-'}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                              <IconButton
                                iconProps={{ iconName: 'Edit' }}
                                title="Edit Quiz"
                                onClick={() => void onEditQuiz(quiz.Id)}
                              />
                              <IconButton
                                iconProps={{ iconName: 'View' }}
                                title="Preview Quiz"
                                onClick={() => void dialogManager.showAlert(`Preview quiz: ${quiz.Title}`, { variant: 'info' })}
                              />
                              <IconButton
                                iconProps={{ iconName: 'BarChartVertical' }}
                                title="View Results"
                                onClick={() => void dialogManager.showAlert(`View results for: ${quiz.Title}`, { variant: 'info' })}
                              />
                            </Stack>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Embedded Quiz Builder Placeholder */}
              <div style={{ marginTop: 24, padding: 24, background: '#f3f2f1', borderRadius: 8, border: '2px dashed #c8c6c4' }}>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
                  <Icon iconName="Frame" style={{ fontSize: 32, color: '#605e5c' }} />
                  <Text variant="medium" style={{ color: '#605e5c' }}>Quiz Editor iframe will be embedded here when editing a quiz</Text>
                </Stack>
              </div>
            </>
          )}
        </div>
      </>
    );
  }
}
