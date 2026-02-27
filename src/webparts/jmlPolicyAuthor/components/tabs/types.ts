// Shared types for extracted PolicyAuthorEnhanced tab components
// These interfaces define the props contracts between the parent component
// and its extracted tab sub-components.

import {
  IPolicyDelegationRequest,
  IPolicyAuthorRequest as IPolicyRequest,
  IAuthorPolicyAnalytics as IPolicyAnalytics,
  IAuthorPolicyQuiz as IPolicyQuiz,
  IAuthorPolicyPack as IPolicyPack,
  IDepartmentCompliance,
  IDelegationKpis,
} from '../../../../models/IPolicyAuthor';

/**
 * Dialog manager interface — subset of createDialogManager() return type
 * used by extracted tab components for user interaction.
 */
export interface IDialogManager {
  showAlert: (message: string, options?: { title?: string; variant?: 'info' | 'success' | 'warning' | 'error' }) => Promise<void>;
  showConfirm: (message: string, options?: { title?: string; confirmText?: string; cancelText?: string }) => Promise<boolean>;
  showPrompt: (message: string, options?: { title?: string; defaultValue?: string; required?: boolean }) => Promise<string | null>;
}

/**
 * Props for the DelegationsTab component
 */
export interface IDelegationsTabProps {
  delegatedRequests: IPolicyDelegationRequest[];
  delegationsLoading: boolean;
  delegationKpis: IDelegationKpis;
  styles: Record<string, string>;
  onNewDelegation: () => void;
  onStartPolicy: (request: IPolicyDelegationRequest) => void;
}

/**
 * Props for the PolicyRequestsTab component
 */
export interface IPolicyRequestsTabProps {
  policyRequests: IPolicyRequest[];
  policyRequestsLoading: boolean;
  requestStatusFilter: string;
  selectedPolicyRequest: IPolicyRequest | null;
  showPolicyRequestDetailPanel: boolean;
  styles: Record<string, string>;
  context: {
    pageContext: {
      user: { displayName: string; email: string };
    };
  };
  onSetState: (stateUpdate: Record<string, unknown>) => void;
  onCreatePolicyFromRequest: (request: IPolicyRequest) => void;
}

/**
 * Props for the AnalyticsTab component
 */
export interface IAnalyticsTabProps {
  analyticsData: IPolicyAnalytics | null;
  analyticsLoading: boolean;
  departmentCompliance: IDepartmentCompliance[];
  styles: Record<string, string>;
  dialogManager: IDialogManager;
  onDateRangeChange: (days: number) => void;
  onExportAnalytics: (format: 'csv' | 'pdf' | 'json') => void;
}

/**
 * Props for the QuizBuilderTab component
 */
export interface IQuizBuilderTabProps {
  quizzes: IPolicyQuiz[];
  quizzesLoading: boolean;
  styles: Record<string, string>;
  dialogManager: IDialogManager;
  onCreateQuiz: () => void;
  onEditQuiz: (quizId: number) => void;
}

/**
 * Props for the PolicyPacksTab component
 */
export interface IPolicyPacksTabProps {
  policyPacks: IPolicyPack[];
  policyPacksLoading: boolean;
  styles: Record<string, string>;
  dialogManager: IDialogManager;
  onCreatePack: () => void;
}
