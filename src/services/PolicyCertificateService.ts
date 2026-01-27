// @ts-nocheck
/**
 * PolicyCertificateService - Stub for standalone Policy Manager
 * PDF generation not available in standalone version
 */

import { logger } from './LoggingService';
import {
  IPolicy,
  IPolicyAcknowledgement,
  IPolicyPack,
  IPolicyPackAssignment,
  IPolicyQuizResult
} from '../models/IPolicy';

// ============================================================================
// INTERFACES
// ============================================================================

export interface ICertificateOptions {
  companyName?: string;
  companyLogoDataUrl?: string;
  showQRCode?: boolean;
  showSignature?: boolean;
  signatureTitle?: string;
  signatureName?: string;
  backgroundColor?: string;
  primaryColor?: string;
  accentColor?: string;
}

export interface ICertificateResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  certificateId?: string;
  error?: string;
}

export interface IAcknowledgementCertificateData {
  certificateId: string;
  employeeName: string;
  employeeEmail: string;
  employeeDepartment?: string;
  employeeRole?: string;
  policyNumber: string;
  policyName: string;
  policyCategory?: string;
  policyVersion?: string;
  acknowledgedDate: Date;
  validUntil?: Date;
  signature?: string;
}

export interface IQuizCertificateData {
  certificateId: string;
  employeeName: string;
  employeeEmail: string;
  employeeDepartment?: string;
  quizTitle: string;
  policyName?: string;
  score: number;
  passingScore: number;
  passed: boolean;
  completedDate: Date;
  totalQuestions: number;
  correctAnswers: number;
  timeSpent?: number;
}

export interface IComplianceCertificateData {
  certificateId: string;
  employeeName: string;
  employeeEmail: string;
  employeeDepartment?: string;
  certificateTitle: string;
  issuedDate: Date;
  validUntil?: Date;
  complianceItems: Array<{
    itemName: string;
    completedDate: Date;
    status: string;
  }>;
}

export interface IBatchCertificateOptions extends ICertificateOptions {
  zipFileName?: string;
  progressCallback?: (current: number, total: number) => void;
}

export interface IBatchCertificateResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  totalGenerated: number;
  errors: Array<{ id: string; error: string }>;
}

// ============================================================================
// SERVICE CLASS (STUB)
// ============================================================================

class PolicyCertificateServiceImpl {
  private static instance: PolicyCertificateServiceImpl;

  private constructor() {}

  public static getInstance(): PolicyCertificateServiceImpl {
    if (!PolicyCertificateServiceImpl.instance) {
      PolicyCertificateServiceImpl.instance = new PolicyCertificateServiceImpl();
    }
    return PolicyCertificateServiceImpl.instance;
  }

  public async generateAcknowledgementCertificate(
    _acknowledgement: IPolicyAcknowledgement,
    _policy: IPolicy,
    _options?: ICertificateOptions
  ): Promise<ICertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      error: 'PDF certificate generation not available in standalone version'
    };
  }

  public async generateQuizCertificate(
    _quizResult: IPolicyQuizResult,
    _options?: ICertificateOptions
  ): Promise<ICertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      error: 'PDF certificate generation not available in standalone version'
    };
  }

  public async generateComplianceCertificate(
    _data: IComplianceCertificateData,
    _options?: ICertificateOptions
  ): Promise<ICertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      error: 'PDF certificate generation not available in standalone version'
    };
  }

  public async generatePolicyPackCertificate(
    _pack: IPolicyPack,
    _assignment: IPolicyPackAssignment,
    _options?: ICertificateOptions
  ): Promise<ICertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      error: 'PDF certificate generation not available in standalone version'
    };
  }

  public async generateBatchAcknowledgementCertificates(
    _acknowledgements: IPolicyAcknowledgement[],
    _policies: Map<number, IPolicy>,
    _options?: IBatchCertificateOptions
  ): Promise<IBatchCertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      totalGenerated: 0,
      errors: [{ id: 'batch', error: 'PDF certificate generation not available in standalone version' }]
    };
  }

  public async generateBatchQuizCertificates(
    _quizResults: IPolicyQuizResult[],
    _options?: IBatchCertificateOptions
  ): Promise<IBatchCertificateResult> {
    logger.warn('PolicyCertificateService', 'PDF generation not available in standalone version');
    return {
      success: false,
      totalGenerated: 0,
      errors: [{ id: 'batch', error: 'PDF certificate generation not available in standalone version' }]
    };
  }

  public generateVerificationUrl(_certificateId: string): string {
    return '';
  }

  public generateCertificateId(): string {
    return `CERT-${Date.now()}-${Math.random().toString(36).substring(2, 8).toUpperCase()}`;
  }
}

export const PolicyCertificateService = PolicyCertificateServiceImpl.getInstance();
export default PolicyCertificateService;
