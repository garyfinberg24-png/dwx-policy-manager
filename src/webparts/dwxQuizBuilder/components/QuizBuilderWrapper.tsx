// @ts-nocheck
import * as React from 'react';
import { IQuizBuilderWrapperProps } from './IQuizBuilderWrapperProps';
import { QuizBuilder } from '../../../components/QuizBuilder';
import { DwxAppLayout } from '../../../components/JmlAppLayout';
import styles from './QuizBuilderWrapper.module.scss';

export const QuizBuilderWrapper: React.FC<IQuizBuilderWrapperProps> = (props) => {
  const { sp, context, title } = props;

  // Get quizId and policyId from URL params if present
  const urlParams = new URLSearchParams(window.location.search);
  const quizId = urlParams.get('quizId') ? parseInt(urlParams.get('quizId')!, 10) : undefined;
  const policyId = urlParams.get('policyId') ? parseInt(urlParams.get('policyId')!, 10) : undefined;

  const handleSave = (quiz: any): void => {
    console.log('Quiz saved:', quiz);
  };

  const handleCancel = (): void => {
    // Navigate back or to quiz list
    window.history.back();
  };

  return (
    <DwxAppLayout
      title={title}
      context={context}
      showBreadcrumb={true}
      fullWidth={true}
    >
      <div className={styles.quizBuilderWrapper}>
        <QuizBuilder
          sp={sp}
          quizId={quizId}
          policyId={policyId}
          onSave={handleSave}
          onCancel={handleCancel}
        />
      </div>
    </DwxAppLayout>
  );
};

export default QuizBuilderWrapper;
