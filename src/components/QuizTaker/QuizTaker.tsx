// @ts-nocheck
/**
 * QuizTaker - Stub component for standalone Policy Manager
 */
import * as React from 'react';

export interface IQuizTakerProps {
  quizId?: number;
  policyId?: number;
  onComplete?: () => void;
  onClose?: () => void;
}

export const QuizTaker: React.FC<IQuizTakerProps> = ({ onClose }) => {
  return (
    <div style={{ padding: '20px', textAlign: 'center' }}>
      <p>Quiz functionality not available in standalone version.</p>
      {onClose && <button onClick={onClose}>Close</button>}
    </div>
  );
};

export default QuizTaker;
