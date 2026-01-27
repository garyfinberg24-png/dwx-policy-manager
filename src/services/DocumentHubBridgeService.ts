// @ts-nocheck
/**
 * Document Hub Bridge Service
 *
 * Provides integration between Document Hub and other JML Enterprise modules:
 * - Contract Manager: Link documents to contracts
 * - Signing Service: Send documents for e-signature
 * - Policy Hub: Apply policies and retention rules
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/files';
import '@pnp/sp/folders';

import {
  IDocumentHubModuleBridge,
  IContractManagerBridge,
  ISigningServiceBridge,
  IPolicyHubBridge,
  ICVManagementBridge,
  IBridgeAvailability,
  IDocumentIntegrationSummary,
  IDocumentContractLink,
  IDocumentSigningRequest,
  IDocumentPolicyLink,
  IDocumentCVLink,
  IPolicyComplianceResult,
  IContractSummary,
  ILinkDocumentToContractRequest,
  DocumentContractLinkType,
  IPolicySummary,
  IApplyPolicyRequest,
  DocumentPolicyLinkType,
  IPolicyRetentionRule,
  IPolicyViolation,
  IPolicyWarning,
  IInitiateSigningRequest,
  IDocumentSigner,
  SigningProvider,
  SigningStatus,
  SignerRole,
  BridgeModuleType,
  PolicyClassification,
  ICVSummary,
  ICVSearchParams,
  ILinkDocumentToCVRequest,
  DocumentCVLinkType
} from '../models/IModuleBridge';
import {
  ICV,
  CVStatus,
  CVSource,
  ExperienceLevel
} from '../models/ICVManagement';
import {
  IDocumentRegistryEntry,
  SourceModule,
  ConfidentialityLevel,
  DocumentStatus
} from '../models/IDocumentHub';
import { logger } from './LoggingService';

// ============================================================================
// LIST NAMES
// ============================================================================

const LISTS = {
  // Document Hub Bridge Lists
  DOCUMENT_CONTRACT_LINKS: 'JML_DocumentContractLinks',
  DOCUMENT_SIGNING_REQUESTS: 'JML_DocumentSigningRequests',
  DOCUMENT_SIGNERS: 'JML_DocumentSigners',
  DOCUMENT_POLICY_LINKS: 'PM_DocumentPolicyLinks',  // Policy Manager list
  DOCUMENT_CV_LINKS: 'JML_DocumentCVLinks',

  // Contract Manager Lists (JML integration - keep JML prefix)
  CONTRACTS: 'JML_ContractRecords',

  // Policy Hub Lists (Policy Manager owned - use PM prefix)
  POLICIES: 'PM_Policies',
  POLICY_CATEGORIES: 'PM_PolicyCategories',
  POLICY_RETENTION: 'PM_PolicyRetention',

  // CV Management Lists (JML integration - keep JML prefix)
  CV_DATABASE: 'JML_CVDatabase',
  CV_POSITIONS: 'JML_Positions',
  CV_DEPARTMENTS: 'JML_Departments',

  // Document Registry (JML integration - keep JML prefix)
  DOCUMENT_REGISTRY: 'JML_DocumentRegistry',

  // Signing Configuration (JML integration - keep JML prefix)
  SIGNATURE_CONFIG: 'JML_SignatureConfig'
};

// ============================================================================
// CONTRACT MANAGER BRIDGE IMPLEMENTATION
// ============================================================================

class ContractManagerBridgeImpl implements IContractManagerBridge {
  private sp: SPFI;
  private currentUserId: number = 0;

  constructor(sp: SPFI, currentUserId: number) {
    this.sp = sp;
    this.currentUserId = currentUserId;
  }

  public async searchContracts(searchText: string, maxResults: number = 20): Promise<IContractSummary[]> {
    try {
      // Check if contracts list exists
      const listExists = await this.listExists(LISTS.CONTRACTS);
      if (!listExists) {
        logger.warn('ContractManagerBridge', 'Contracts list does not exist');
        return [];
      }

      const items = await this.sp.web.lists
        .getByTitle(LISTS.CONTRACTS)
        .items
        .filter(`substringof('${searchText}', Title) or substringof('${searchText}', ContractNumber)`)
        .select('Id', 'Title', 'ContractNumber', 'Status', 'PartyName', 'StartDate', 'EndDate', 'TotalValue', 'Currency')
        .top(maxResults)();

      return items.map((item: any) => ({
        id: item.Id,
        contractNumber: item.ContractNumber || '',
        title: item.Title,
        status: item.Status || 'Draft',
        partyName: item.PartyName || '',
        startDate: item.StartDate ? new Date(item.StartDate) : undefined,
        endDate: item.EndDate ? new Date(item.EndDate) : undefined,
        value: item.TotalValue,
        currency: item.Currency
      }));
    } catch (error) {
      logger.error('ContractManagerBridge', 'Error searching contracts:', error);
      return [];
    }
  }

  public async linkDocumentToContract(request: ILinkDocumentToContractRequest): Promise<IDocumentContractLink> {
    try {
      // Get contract details
      const contract = await this.sp.web.lists
        .getByTitle(LISTS.CONTRACTS)
        .items
        .getById(request.contractId)
        .select('Id', 'Title', 'ContractNumber')();

      // Create link
      const result = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CONTRACT_LINKS)
        .items
        .add({
          Title: `Doc-${request.documentId}-Contract-${request.contractId}`,
          DocumentId: request.documentId,
          ContractId: request.contractId,
          ContractNumber: contract.ContractNumber,
          ContractTitle: contract.Title,
          LinkType: request.linkType,
          LinkedDate: new Date().toISOString(),
          LinkedById: this.currentUserId,
          Notes: request.notes || ''
        });

      return {
        id: result.data.Id,
        documentId: request.documentId,
        contractId: request.contractId,
        contractNumber: contract.ContractNumber,
        contractTitle: contract.Title,
        linkType: request.linkType,
        linkedDate: new Date(),
        linkedBy: '',
        notes: request.notes
      };
    } catch (error) {
      logger.error('ContractManagerBridge', 'Error linking document to contract:', error);
      throw error;
    }
  }

  public async unlinkDocumentFromContract(linkId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CONTRACT_LINKS)
        .items
        .getById(linkId)
        .delete();
    } catch (error) {
      logger.error('ContractManagerBridge', 'Error unlinking document from contract:', error);
      throw error;
    }
  }

  public async getLinkedContracts(documentId: number): Promise<IDocumentContractLink[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_CONTRACT_LINKS);
      if (!listExists) return [];

      const items = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CONTRACT_LINKS)
        .items
        .filter(`DocumentId eq ${documentId}`)
        .select('*', 'LinkedBy/Title')
        .expand('LinkedBy')();

      return items.map((item: any) => ({
        id: item.Id,
        documentId: item.DocumentId,
        contractId: item.ContractId,
        contractNumber: item.ContractNumber || '',
        contractTitle: item.ContractTitle || '',
        linkType: item.LinkType as DocumentContractLinkType,
        linkedDate: new Date(item.LinkedDate),
        linkedBy: item.LinkedBy?.Title || 'Unknown',
        notes: item.Notes
      }));
    } catch (error) {
      logger.error('ContractManagerBridge', 'Error getting linked contracts:', error);
      return [];
    }
  }

  public async getLinkedDocuments(contractId: number): Promise<IDocumentRegistryEntry[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_CONTRACT_LINKS);
      if (!listExists) return [];

      // Get document IDs linked to contract
      const links = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CONTRACT_LINKS)
        .items
        .filter(`ContractId eq ${contractId}`)
        .select('DocumentId')();

      if (links.length === 0) return [];

      // Get document details
      const documentIds = links.map((l: any) => l.DocumentId);
      const filterQuery = documentIds.map((id: number) => `Id eq ${id}`).join(' or ');

      const documents = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .filter(filterQuery)
        .select('*')();

      return documents.map((doc: any) => this.mapToDocumentRegistryEntry(doc));
    } catch (error) {
      logger.error('ContractManagerBridge', 'Error getting linked documents:', error);
      return [];
    }
  }

  private async listExists(listName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Title')();
      return true;
    } catch {
      return false;
    }
  }

  private mapToDocumentRegistryEntry(item: any): IDocumentRegistryEntry {
    return {
      Id: item.Id,
      Title: item.Title,
      DocumentId: item.DocumentId || '',
      SourceModule: (item.SourceModule as SourceModule) || SourceModule.ManualUpload,
      SourceItemId: item.SourceItemId,
      SourceUrl: item.SourceUrl || item.FileUrl,
      ConfidentialityLevel: (item.ConfidentialityLevel as ConfidentialityLevel) || ConfidentialityLevel.Internal,
      DocumentStatus: (item.DocumentStatus || item.Status || DocumentStatus.Active) as DocumentStatus,
      IsRecord: item.IsRecord || false,
      OnLegalHold: item.OnLegalHold || false,
      VersionCount: item.VersionCount || 1,
      ExternalAccessEnabled: item.ExternalAccessEnabled || false,
      ActiveShareCount: item.ActiveShareCount || 0,
      ViewCount: item.ViewCount || 0,
      DownloadCount: item.DownloadCount || 0,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}

// ============================================================================
// SIGNING SERVICE BRIDGE IMPLEMENTATION
// ============================================================================

class SigningServiceBridgeImpl implements ISigningServiceBridge {
  private sp: SPFI;
  private currentUserId: number = 0;

  constructor(sp: SPFI, currentUserId: number) {
    this.sp = sp;
    this.currentUserId = currentUserId;
  }

  public async getAvailableProviders(): Promise<SigningProvider[]> {
    try {
      const listExists = await this.listExists(LISTS.SIGNATURE_CONFIG);
      if (!listExists) {
        // Default to internal signing if no config exists
        return [SigningProvider.Internal];
      }

      const configs = await this.sp.web.lists
        .getByTitle(LISTS.SIGNATURE_CONFIG)
        .items
        .filter("IsEnabled eq 1")
        .select('Provider')();

      const providers = configs.map((c: any) => c.Provider as SigningProvider);

      // Always include internal provider
      if (!providers.includes(SigningProvider.Internal)) {
        providers.unshift(SigningProvider.Internal);
      }

      return providers;
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error getting available providers:', error);
      return [SigningProvider.Internal];
    }
  }

  public async canSignDocument(documentId: number): Promise<{ canSign: boolean; reason?: string }> {
    try {
      // Get document details
      const document = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .getById(documentId)
        .select('Status', 'FileUrl', 'Classification')();

      // Check if document status allows signing
      if (document.Status === 'Archived' || document.Status === 'Disposed') {
        return { canSign: false, reason: 'Document is archived or disposed' };
      }

      // Check if document has a file
      if (!document.FileUrl) {
        return { canSign: false, reason: 'Document has no associated file' };
      }

      // Check for existing active signing request
      const existingRequest = await this.getSigningStatus(documentId);
      if (existingRequest && [SigningStatus.Pending, SigningStatus.InProgress].includes(existingRequest.status)) {
        return { canSign: false, reason: 'Document already has an active signing request' };
      }

      return { canSign: true };
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error checking if document can be signed:', error);
      return { canSign: false, reason: 'Error checking document status' };
    }
  }

  public async initiateSigningRequest(request: IInitiateSigningRequest): Promise<IDocumentSigningRequest> {
    try {
      // Create signing request
      const signingResult = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .add({
          Title: `Signing-${request.documentId}-${new Date().getTime()}`,
          DocumentId: request.documentId,
          Provider: request.provider,
          Status: SigningStatus.Pending,
          EmailSubject: request.emailSubject,
          EmailMessage: request.emailMessage || '',
          ExpirationDays: request.expirationDays || 30,
          ReminderDays: request.reminderDays || 3,
          CreatedById: this.currentUserId
        });

      const requestId = signingResult.data.Id;

      // Create signer records
      const signers: IDocumentSigner[] = [];
      for (const signer of request.signers) {
        const signerResult = await this.sp.web.lists
          .getByTitle(LISTS.DOCUMENT_SIGNERS)
          .items
          .add({
            Title: signer.name,
            SigningRequestId: requestId,
            SignerName: signer.name,
            SignerEmail: signer.email,
            SignerRole: signer.role,
            SignerOrder: signer.order,
            Status: SigningStatus.Pending
          });

        signers.push({
          id: signerResult.data.Id,
          name: signer.name,
          email: signer.email,
          role: signer.role,
          order: signer.order,
          status: SigningStatus.Pending
        });
      }

      // If using external provider, initiate with that service
      if (request.provider !== SigningProvider.Internal) {
        // TODO: Call external signing service (DocuSign/AdobeSign)
        // For now, just log that this would happen
        logger.info('SigningServiceBridge', `Would initiate ${request.provider} signing for document ${request.documentId}`);
      }

      return {
        id: requestId,
        documentId: request.documentId,
        provider: request.provider,
        status: SigningStatus.Pending,
        signers: signers,
        emailSubject: request.emailSubject,
        emailMessage: request.emailMessage,
        expirationDays: request.expirationDays,
        reminderDays: request.reminderDays,
        createdDate: new Date()
      };
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error initiating signing request:', error);
      throw error;
    }
  }

  public async getSigningStatus(documentId: number): Promise<IDocumentSigningRequest | null> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_SIGNING_REQUESTS);
      if (!listExists) return null;

      const requests = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .filter(`DocumentId eq ${documentId}`)
        .orderBy('Created', false)
        .top(1)
        .select('*', 'CreatedBy/Title')
        .expand('CreatedBy')();

      if (requests.length === 0) return null;

      const request = requests[0];
      const signers = await this.getSigners(request.Id);

      return {
        id: request.Id,
        documentId: request.DocumentId,
        provider: request.Provider as SigningProvider,
        status: request.Status as SigningStatus,
        signers: signers,
        emailSubject: request.EmailSubject,
        emailMessage: request.EmailMessage,
        expirationDays: request.ExpirationDays,
        reminderDays: request.ReminderDays,
        createdDate: new Date(request.Created),
        createdBy: request.CreatedBy?.Title,
        completedDate: request.CompletedDate ? new Date(request.CompletedDate) : undefined,
        externalEnvelopeId: request.ExternalEnvelopeId
      };
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error getting signing status:', error);
      return null;
    }
  }

  public async getSigningHistory(documentId: number): Promise<IDocumentSigningRequest[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_SIGNING_REQUESTS);
      if (!listExists) return [];

      const requests = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .filter(`DocumentId eq ${documentId}`)
        .orderBy('Created', false)
        .select('*', 'CreatedBy/Title')
        .expand('CreatedBy')();

      const result: IDocumentSigningRequest[] = [];
      for (const request of requests) {
        const signers = await this.getSigners(request.Id);
        result.push({
          id: request.Id,
          documentId: request.DocumentId,
          provider: request.Provider as SigningProvider,
          status: request.Status as SigningStatus,
          signers: signers,
          emailSubject: request.EmailSubject,
          emailMessage: request.EmailMessage,
          expirationDays: request.ExpirationDays,
          reminderDays: request.ReminderDays,
          createdDate: new Date(request.Created),
          createdBy: request.CreatedBy?.Title,
          completedDate: request.CompletedDate ? new Date(request.CompletedDate) : undefined,
          externalEnvelopeId: request.ExternalEnvelopeId
        });
      }

      return result;
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error getting signing history:', error);
      return [];
    }
  }

  public async cancelSigningRequest(requestId: number, reason: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .getById(requestId)
        .update({
          Status: SigningStatus.Voided,
          VoidReason: reason,
          VoidedDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error canceling signing request:', error);
      throw error;
    }
  }

  public async resendSigningRequest(requestId: number): Promise<void> {
    try {
      // Get the request details
      const request = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .getById(requestId)
        .select('*')();

      // Update last sent date
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .getById(requestId)
        .update({
          LastSentDate: new Date().toISOString()
        });

      // If external provider, call their resend API
      if (request.Provider !== SigningProvider.Internal && request.ExternalEnvelopeId) {
        // TODO: Call external provider's resend API
        logger.info('SigningServiceBridge', `Would resend ${request.Provider} request ${request.ExternalEnvelopeId}`);
      }
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error resending signing request:', error);
      throw error;
    }
  }

  public async downloadSignedDocument(requestId: number): Promise<Blob> {
    try {
      const request = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNING_REQUESTS)
        .items
        .getById(requestId)
        .select('SignedDocumentUrl')();

      if (!request.SignedDocumentUrl) {
        throw new Error('Signed document not available');
      }

      // Download the file
      const fileBlob = await this.sp.web.getFileByServerRelativePath(request.SignedDocumentUrl).getBlob();
      return fileBlob;
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error downloading signed document:', error);
      throw error;
    }
  }

  private async getSigners(requestId: number): Promise<IDocumentSigner[]> {
    try {
      const signers = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_SIGNERS)
        .items
        .filter(`SigningRequestId eq ${requestId}`)
        .orderBy('SignerOrder')
        .select('*')();

      return signers.map((s: any) => ({
        id: s.Id,
        name: s.SignerName,
        email: s.SignerEmail,
        role: s.SignerRole as SignerRole,
        order: s.SignerOrder,
        status: s.Status as SigningStatus,
        signedDate: s.SignedDate ? new Date(s.SignedDate) : undefined,
        declineReason: s.DeclineReason
      }));
    } catch (error) {
      logger.error('SigningServiceBridge', 'Error getting signers:', error);
      return [];
    }
  }

  private async listExists(listName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Title')();
      return true;
    } catch {
      return false;
    }
  }
}

// ============================================================================
// POLICY HUB BRIDGE IMPLEMENTATION
// ============================================================================

class PolicyHubBridgeImpl implements IPolicyHubBridge {
  private sp: SPFI;
  private currentUserId: number = 0;

  constructor(sp: SPFI, currentUserId: number) {
    this.sp = sp;
    this.currentUserId = currentUserId;
  }

  public async searchPolicies(searchText: string, category?: string, maxResults: number = 20): Promise<IPolicySummary[]> {
    try {
      const listExists = await this.listExists(LISTS.POLICIES);
      if (!listExists) {
        logger.warn('PolicyHubBridge', 'Policies list does not exist');
        return [];
      }

      let filterQuery = `substringof('${searchText}', Title) or substringof('${searchText}', PolicyNumber)`;
      if (category) {
        filterQuery = `(${filterQuery}) and Category eq '${category}'`;
      }

      const items = await this.sp.web.lists
        .getByTitle(LISTS.POLICIES)
        .items
        .filter(filterQuery)
        .select('Id', 'Title', 'PolicyNumber', 'Category', 'Classification', 'EffectiveDate', 'ReviewDate', 'Status', 'Owner/Title')
        .expand('Owner')
        .top(maxResults)();

      return items.map((item: any) => ({
        id: item.Id,
        policyNumber: item.PolicyNumber || '',
        title: item.Title,
        category: item.Category || 'General',
        classification: item.Classification as PolicyClassification || PolicyClassification.Internal,
        effectiveDate: new Date(item.EffectiveDate || new Date()),
        reviewDate: item.ReviewDate ? new Date(item.ReviewDate) : undefined,
        owner: item.Owner?.Title || 'Unassigned',
        status: item.Status || 'Draft'
      }));
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error searching policies:', error);
      return [];
    }
  }

  public async getPolicyCategories(): Promise<string[]> {
    try {
      const listExists = await this.listExists(LISTS.POLICY_CATEGORIES);
      if (!listExists) {
        // Return default categories
        return ['HR', 'IT', 'Finance', 'Legal', 'Operations', 'Safety', 'General'];
      }

      const categories = await this.sp.web.lists
        .getByTitle(LISTS.POLICY_CATEGORIES)
        .items
        .select('Title')
        .orderBy('Title')();

      return categories.map((c: any) => c.Title);
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error getting policy categories:', error);
      return ['HR', 'IT', 'Finance', 'Legal', 'Operations', 'Safety', 'General'];
    }
  }

  public async linkDocumentToPolicy(request: IApplyPolicyRequest): Promise<IDocumentPolicyLink> {
    try {
      // Get policy details
      const policy = await this.sp.web.lists
        .getByTitle(LISTS.POLICIES)
        .items
        .getById(request.policyId)
        .select('Id', 'Title', 'PolicyNumber', 'Classification')();

      // Create link
      const result = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_POLICY_LINKS)
        .items
        .add({
          Title: `Doc-${request.documentId}-Policy-${request.policyId}`,
          DocumentId: request.documentId,
          PolicyId: request.policyId,
          PolicyNumber: policy.PolicyNumber,
          PolicyTitle: policy.Title,
          LinkType: request.linkType,
          LinkedDate: new Date().toISOString(),
          LinkedById: this.currentUserId,
          Notes: request.notes || ''
        });

      // Apply retention if requested
      if (request.applyRetention) {
        await this.applyPolicyRetention(request.documentId, request.policyId);
      }

      // Apply classification if requested
      if (request.applyClassification) {
        await this.applyPolicyClassification(request.documentId, request.policyId);
      }

      return {
        id: result.data.Id,
        documentId: request.documentId,
        policyId: request.policyId,
        policyNumber: policy.PolicyNumber,
        policyTitle: policy.Title,
        linkType: request.linkType,
        linkedDate: new Date(),
        linkedBy: '',
        notes: request.notes
      };
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error linking document to policy:', error);
      throw error;
    }
  }

  public async unlinkDocumentFromPolicy(linkId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_POLICY_LINKS)
        .items
        .getById(linkId)
        .delete();
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error unlinking document from policy:', error);
      throw error;
    }
  }

  public async getLinkedPolicies(documentId: number): Promise<IDocumentPolicyLink[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_POLICY_LINKS);
      if (!listExists) return [];

      const items = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_POLICY_LINKS)
        .items
        .filter(`DocumentId eq ${documentId}`)
        .select('*', 'LinkedBy/Title')
        .expand('LinkedBy')();

      return items.map((item: any) => ({
        id: item.Id,
        documentId: item.DocumentId,
        policyId: item.PolicyId,
        policyNumber: item.PolicyNumber || '',
        policyTitle: item.PolicyTitle || '',
        linkType: item.LinkType as DocumentPolicyLinkType,
        linkedDate: new Date(item.LinkedDate),
        linkedBy: item.LinkedBy?.Title || 'Unknown',
        notes: item.Notes
      }));
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error getting linked policies:', error);
      return [];
    }
  }

  public async getPolicyRetentionRule(policyId: number): Promise<IPolicyRetentionRule | null> {
    try {
      const listExists = await this.listExists(LISTS.POLICY_RETENTION);
      if (!listExists) return null;

      const rules = await this.sp.web.lists
        .getByTitle(LISTS.POLICY_RETENTION)
        .items
        .filter(`PolicyId eq ${policyId}`)
        .top(1)
        .select('*')();

      if (rules.length === 0) return null;

      const rule = rules[0];
      return {
        id: rule.Id,
        policyId: rule.PolicyId,
        retentionPeriod: rule.RetentionPeriod || 7,
        retentionUnit: rule.RetentionUnit || 'Years',
        dispositionAction: rule.DispositionAction || 'Archive',
        triggerEvent: rule.TriggerEvent || 'Creation',
        customTriggerDate: rule.CustomTriggerDate ? new Date(rule.CustomTriggerDate) : undefined
      };
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error getting policy retention rule:', error);
      return null;
    }
  }

  public async applyPolicyClassification(documentId: number, policyId: number): Promise<void> {
    try {
      // Get policy classification
      const policy = await this.sp.web.lists
        .getByTitle(LISTS.POLICIES)
        .items
        .getById(policyId)
        .select('Classification')();

      // Map policy classification to document classification
      let documentClassification = 'Internal';
      switch (policy.Classification) {
        case 'Public':
          documentClassification = 'Public';
          break;
        case 'Confidential':
          documentClassification = 'Confidential';
          break;
        case 'Restricted':
          documentClassification = 'Restricted';
          break;
        default:
          documentClassification = 'Internal';
      }

      // Update document classification
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .getById(documentId)
        .update({
          Classification: documentClassification,
          ClassificationSource: 'Policy',
          ClassificationPolicyId: policyId
        });
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error applying policy classification:', error);
      throw error;
    }
  }

  public async applyPolicyRetention(documentId: number, policyId: number): Promise<void> {
    try {
      // Get policy retention rule
      const rule = await this.getPolicyRetentionRule(policyId);
      if (!rule) {
        logger.warn('PolicyHubBridge', `No retention rule found for policy ${policyId}`);
        return;
      }

      // Calculate retention end date
      const document = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .getById(documentId)
        .select('Created', 'Modified')();

      let triggerDate = new Date();
      switch (rule.triggerEvent) {
        case 'Creation':
          triggerDate = new Date(document.Created);
          break;
        case 'LastModified':
          triggerDate = new Date(document.Modified);
          break;
        case 'Custom':
          triggerDate = rule.customTriggerDate || new Date();
          break;
        default:
          triggerDate = new Date();
      }

      // Calculate end date
      const retentionEndDate = new Date(triggerDate);
      switch (rule.retentionUnit) {
        case 'Days':
          retentionEndDate.setDate(retentionEndDate.getDate() + rule.retentionPeriod);
          break;
        case 'Months':
          retentionEndDate.setMonth(retentionEndDate.getMonth() + rule.retentionPeriod);
          break;
        case 'Years':
          retentionEndDate.setFullYear(retentionEndDate.getFullYear() + rule.retentionPeriod);
          break;
      }

      // Update document retention
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .getById(documentId)
        .update({
          RetentionPolicyId: policyId,
          RetentionStartDate: triggerDate.toISOString(),
          RetentionEndDate: retentionEndDate.toISOString(),
          DispositionAction: rule.dispositionAction
        });
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error applying policy retention:', error);
      throw error;
    }
  }

  public async checkPolicyCompliance(documentId: number): Promise<IPolicyComplianceResult> {
    const violations: IPolicyViolation[] = [];
    const warnings: IPolicyWarning[] = [];

    try {
      // Get linked policies
      const linkedPolicies = await this.getLinkedPolicies(documentId);

      // Get document details
      const document = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .getById(documentId)
        .select('*')();

      // Check each linked policy
      for (const link of linkedPolicies) {
        // Get policy details
        const policy = await this.sp.web.lists
          .getByTitle(LISTS.POLICIES)
          .items
          .getById(link.policyId)
          .select('*')();

        // Check classification compliance
        if (policy.Classification && document.Classification !== policy.Classification) {
          violations.push({
            policyId: link.policyId,
            policyTitle: link.policyTitle,
            requirement: `Document classification must match policy classification (${policy.Classification})`,
            violationType: 'Classification',
            severity: 'High',
            remediation: `Update document classification to ${policy.Classification}`
          });
        }

        // Check retention compliance
        if (policy.RequiresRetention && !document.RetentionPolicyId) {
          violations.push({
            policyId: link.policyId,
            policyTitle: link.policyTitle,
            requirement: 'Document must have retention policy applied',
            violationType: 'Retention',
            severity: 'Medium',
            remediation: 'Apply retention settings from linked policy'
          });
        }

        // Check for upcoming review
        if (policy.ReviewDate) {
          const reviewDate = new Date(policy.ReviewDate);
          const now = new Date();
          const daysUntilReview = Math.ceil((reviewDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

          if (daysUntilReview <= 30 && daysUntilReview > 0) {
            warnings.push({
              policyId: link.policyId,
              policyTitle: link.policyTitle,
              message: `Policy review due in ${daysUntilReview} days`,
              warningType: 'UpcomingReview'
            });
          }
        }

        // Check retention expiration
        if (document.RetentionEndDate) {
          const retentionEnd = new Date(document.RetentionEndDate);
          const now = new Date();
          const daysUntilExpiry = Math.ceil((retentionEnd.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

          if (daysUntilExpiry <= 90 && daysUntilExpiry > 0) {
            warnings.push({
              policyId: link.policyId,
              policyTitle: link.policyTitle,
              message: `Document retention expires in ${daysUntilExpiry} days`,
              warningType: 'RetentionExpiring'
            });
          }
        }
      }

      return {
        documentId: documentId,
        isCompliant: violations.length === 0,
        checkedDate: new Date(),
        linkedPolicies: linkedPolicies.length,
        violations: violations,
        warnings: warnings
      };
    } catch (error) {
      logger.error('PolicyHubBridge', 'Error checking policy compliance:', error);
      return {
        documentId: documentId,
        isCompliant: false,
        checkedDate: new Date(),
        linkedPolicies: 0,
        violations: [{
          policyId: 0,
          policyTitle: 'System',
          requirement: 'Compliance check failed',
          violationType: 'Other',
          severity: 'Low',
          remediation: 'Please try again or contact support'
        }],
        warnings: []
      };
    }
  }

  private async listExists(listName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Title')();
      return true;
    } catch {
      return false;
    }
  }
}

// ============================================================================
// CV MANAGEMENT BRIDGE IMPLEMENTATION
// ============================================================================

class CVManagementBridgeImpl implements ICVManagementBridge {
  private sp: SPFI;
  private currentUserId: number = 0;

  constructor(sp: SPFI, currentUserId: number) {
    this.sp = sp;
    this.currentUserId = currentUserId;
  }

  public async searchCVs(params: ICVSearchParams): Promise<ICVSummary[]> {
    try {
      const listExists = await this.listExists(LISTS.CV_DATABASE);
      if (!listExists) {
        logger.warn('CVManagementBridge', 'CV Database list does not exist');
        return [];
      }

      // Build filter query
      const filters: string[] = [];

      if (params.keyword) {
        filters.push(`(substringof('${params.keyword}', CandidateName) or substringof('${params.keyword}', Email) or substringof('${params.keyword}', PositionAppliedFor) or substringof('${params.keyword}', Skills))`);
      }

      if (params.candidateName) {
        filters.push(`substringof('${params.candidateName}', CandidateName)`);
      }

      if (params.email) {
        filters.push(`substringof('${params.email}', Email)`);
      }

      if (params.positionAppliedFor) {
        filters.push(`PositionAppliedFor eq '${params.positionAppliedFor}'`);
      }

      if (params.department) {
        filters.push(`Department eq '${params.department}'`);
      }

      if (params.status && params.status.length > 0) {
        const statusFilters = params.status.map(s => `Status eq '${s}'`).join(' or ');
        filters.push(`(${statusFilters})`);
      }

      if (params.source && params.source.length > 0) {
        const sourceFilters = params.source.map(s => `Source eq '${s}'`).join(' or ');
        filters.push(`(${sourceFilters})`);
      }

      if (params.experienceLevel && params.experienceLevel.length > 0) {
        const expFilters = params.experienceLevel.map(e => `ExperienceLevel eq '${e}'`).join(' or ');
        filters.push(`(${expFilters})`);
      }

      if (params.minQualificationScore !== undefined) {
        filters.push(`QualificationScore ge ${params.minQualificationScore}`);
      }

      if (params.submissionDateFrom) {
        filters.push(`SubmissionDate ge datetime'${params.submissionDateFrom.toISOString()}'`);
      }

      if (params.submissionDateTo) {
        filters.push(`SubmissionDate le datetime'${params.submissionDateTo.toISOString()}'`);
      }

      const filterQuery = filters.length > 0 ? filters.join(' and ') : '';
      const maxResults = params.maxResults || 50;

      let query = this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .select('Id', 'Title', 'CandidateName', 'Email', 'PositionAppliedFor', 'Department',
                'Status', 'Source', 'SubmissionDate', 'QualificationScore', 'ExperienceLevel',
                'Skills', 'CVFileUrl', 'CVFileName')
        .top(maxResults)
        .orderBy('SubmissionDate', false);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();

      return items.map((item: any) => this.mapToCVSummary(item));
    } catch (error) {
      logger.error('CVManagementBridge', 'Error searching CVs:', error);
      return [];
    }
  }

  public async getCVById(cvId: number): Promise<ICV | null> {
    try {
      const listExists = await this.listExists(LISTS.CV_DATABASE);
      if (!listExists) return null;

      const item = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .getById(cvId)
        .select('*', 'Reviewer/Title', 'Reviewer/EMail', 'ShortlistedBy/Title', 'RejectedBy/Title')
        .expand('Reviewer', 'ShortlistedBy', 'RejectedBy')();

      return this.mapToCV(item);
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting CV by ID:', error);
      return null;
    }
  }

  public async getCVSummary(cvId: number): Promise<ICVSummary | null> {
    try {
      const listExists = await this.listExists(LISTS.CV_DATABASE);
      if (!listExists) return null;

      const item = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .getById(cvId)
        .select('Id', 'Title', 'CandidateName', 'Email', 'PositionAppliedFor', 'Department',
                'Status', 'Source', 'SubmissionDate', 'QualificationScore', 'ExperienceLevel',
                'Skills', 'CVFileUrl', 'CVFileName')();

      return this.mapToCVSummary(item);
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting CV summary:', error);
      return null;
    }
  }

  public async linkDocumentToCV(request: ILinkDocumentToCVRequest): Promise<IDocumentCVLink> {
    try {
      // Get CV details
      const cv = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .getById(request.cvId)
        .select('Id', 'CandidateName', 'Email')();

      // Create link
      const result = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CV_LINKS)
        .items
        .add({
          Title: `Doc-${request.documentId}-CV-${request.cvId}`,
          DocumentId: request.documentId,
          CVId: request.cvId,
          CandidateName: cv.CandidateName,
          CandidateEmail: cv.Email,
          LinkType: request.linkType,
          LinkedDate: new Date().toISOString(),
          LinkedById: this.currentUserId,
          Notes: request.notes || ''
        });

      return {
        id: result.data.Id,
        documentId: request.documentId,
        cvId: request.cvId,
        candidateName: cv.CandidateName,
        candidateEmail: cv.Email,
        linkType: request.linkType,
        linkedDate: new Date(),
        linkedBy: '',
        notes: request.notes
      };
    } catch (error) {
      logger.error('CVManagementBridge', 'Error linking document to CV:', error);
      throw error;
    }
  }

  public async unlinkDocumentFromCV(linkId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CV_LINKS)
        .items
        .getById(linkId)
        .delete();
    } catch (error) {
      logger.error('CVManagementBridge', 'Error unlinking document from CV:', error);
      throw error;
    }
  }

  public async getLinkedCVs(documentId: number): Promise<IDocumentCVLink[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_CV_LINKS);
      if (!listExists) return [];

      const items = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CV_LINKS)
        .items
        .filter(`DocumentId eq ${documentId}`)
        .select('*', 'LinkedBy/Title')
        .expand('LinkedBy')();

      return items.map((item: any) => ({
        id: item.Id,
        documentId: item.DocumentId,
        cvId: item.CVId,
        candidateName: item.CandidateName || '',
        candidateEmail: item.CandidateEmail || '',
        linkType: item.LinkType as DocumentCVLinkType,
        linkedDate: new Date(item.LinkedDate),
        linkedBy: item.LinkedBy?.Title || 'Unknown',
        notes: item.Notes
      }));
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting linked CVs:', error);
      return [];
    }
  }

  public async getLinkedDocuments(cvId: number): Promise<IDocumentRegistryEntry[]> {
    try {
      const listExists = await this.listExists(LISTS.DOCUMENT_CV_LINKS);
      if (!listExists) return [];

      // Get document IDs linked to CV
      const links = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_CV_LINKS)
        .items
        .filter(`CVId eq ${cvId}`)
        .select('DocumentId')();

      if (links.length === 0) return [];

      // Get document details
      const documentIds = links.map((l: any) => l.DocumentId);
      const filterQuery = documentIds.map((id: number) => `Id eq ${id}`).join(' or ');

      const documents = await this.sp.web.lists
        .getByTitle(LISTS.DOCUMENT_REGISTRY)
        .items
        .filter(filterQuery)
        .select('*')();

      return documents.map((doc: any) => this.mapToDocumentRegistryEntry(doc));
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting linked documents:', error);
      return [];
    }
  }

  public async getCVDocument(cvId: number): Promise<{ url: string; fileName: string } | null> {
    try {
      const cv = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .getById(cvId)
        .select('CVFileUrl', 'CVFileName')();

      if (!cv.CVFileUrl) return null;

      return {
        url: cv.CVFileUrl,
        fileName: cv.CVFileName || 'CV.pdf'
      };
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting CV document:', error);
      return null;
    }
  }

  public async getPositions(): Promise<string[]> {
    try {
      // Try to get from positions list first
      const listExists = await this.listExists(LISTS.CV_POSITIONS);
      if (listExists) {
        const positions = await this.sp.web.lists
          .getByTitle(LISTS.CV_POSITIONS)
          .items
          .select('Title')
          .orderBy('Title')();
        return positions.map((p: any) => p.Title);
      }

      // Fallback: get unique positions from CV database
      const cvListExists = await this.listExists(LISTS.CV_DATABASE);
      if (!cvListExists) return [];

      const cvs = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .select('PositionAppliedFor')
        .filter("PositionAppliedFor ne null")();

      const uniquePositions = Array.from(new Set(cvs.map((c: any) => c.PositionAppliedFor).filter(Boolean)));
      return uniquePositions.sort();
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting positions:', error);
      return [];
    }
  }

  public async getDepartments(): Promise<string[]> {
    try {
      // Try to get from departments list first
      const listExists = await this.listExists(LISTS.CV_DEPARTMENTS);
      if (listExists) {
        const departments = await this.sp.web.lists
          .getByTitle(LISTS.CV_DEPARTMENTS)
          .items
          .select('Title')
          .orderBy('Title')();
        return departments.map((d: any) => d.Title);
      }

      // Fallback: get unique departments from CV database
      const cvListExists = await this.listExists(LISTS.CV_DATABASE);
      if (!cvListExists) return [];

      const cvs = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .select('Department')
        .filter("Department ne null")();

      const uniqueDepartments = Array.from(new Set(cvs.map((c: any) => c.Department).filter(Boolean)));
      return uniqueDepartments.sort();
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting departments:', error);
      return [];
    }
  }

  public async getAvailableSkills(): Promise<string[]> {
    try {
      const listExists = await this.listExists(LISTS.CV_DATABASE);
      if (!listExists) return [];

      const cvs = await this.sp.web.lists
        .getByTitle(LISTS.CV_DATABASE)
        .items
        .select('Skills')
        .filter("Skills ne null")
        .top(500)();

      // Extract and flatten skills (stored as comma-separated or JSON array)
      const allSkills: string[] = [];
      for (const cv of cvs) {
        if (cv.Skills) {
          try {
            // Try parsing as JSON array first
            const skills = JSON.parse(cv.Skills);
            if (Array.isArray(skills)) {
              allSkills.push(...skills);
            }
          } catch {
            // Fallback: treat as comma-separated string
            const skills = cv.Skills.split(',').map((s: string) => s.trim());
            allSkills.push(...skills);
          }
        }
      }

      // Get unique skills and sort
      const uniqueSkills = Array.from(new Set(allSkills.filter(Boolean)));
      return uniqueSkills.sort();
    } catch (error) {
      logger.error('CVManagementBridge', 'Error getting available skills:', error);
      return [];
    }
  }

  private mapToCVSummary(item: any): ICVSummary {
    let skills: string[] = [];
    if (item.Skills) {
      try {
        skills = JSON.parse(item.Skills);
      } catch {
        skills = item.Skills.split(',').map((s: string) => s.trim());
      }
    }

    return {
      id: item.Id,
      candidateName: item.CandidateName || item.Title,
      email: item.Email || '',
      positionAppliedFor: item.PositionAppliedFor,
      department: item.Department,
      status: (item.Status as CVStatus) || CVStatus.New,
      source: (item.Source as CVSource) || CVSource.DirectApplication,
      submissionDate: new Date(item.SubmissionDate || item.Created),
      qualificationScore: item.QualificationScore,
      experienceLevel: item.ExperienceLevel as ExperienceLevel,
      skills: skills,
      cvFileUrl: item.CVFileUrl,
      cvFileName: item.CVFileName
    };
  }

  private mapToCV(item: any): ICV {
    let skills: string[] = [];
    let keywordTags: string[] = [];

    if (item.Skills) {
      try {
        skills = JSON.parse(item.Skills);
      } catch {
        skills = item.Skills.split(',').map((s: string) => s.trim());
      }
    }

    if (item.KeywordTags) {
      try {
        keywordTags = JSON.parse(item.KeywordTags);
      } catch {
        keywordTags = item.KeywordTags.split(',').map((s: string) => s.trim());
      }
    }

    return {
      Id: item.Id,
      Title: item.Title,
      CandidateName: item.CandidateName || item.Title,
      Email: item.Email || '',
      Phone: item.Phone,
      LinkedInProfile: item.LinkedInProfile,
      Location: item.Location,
      CVFileName: item.CVFileName || '',
      CVFileUrl: item.CVFileUrl,
      CVFileSize: item.CVFileSize,
      SubmissionDate: new Date(item.SubmissionDate || item.Created),
      Source: (item.Source as CVSource) || CVSource.DirectApplication,
      PositionAppliedFor: item.PositionAppliedFor,
      JobRequisitionId: item.JobRequisitionId,
      Department: item.Department,
      YearsOfExperience: item.YearsOfExperience,
      ExperienceLevel: item.ExperienceLevel as ExperienceLevel,
      HighestEducation: item.HighestEducation,
      Skills: skills,
      Certifications: item.Certifications,
      Languages: item.Languages,
      Status: (item.Status as CVStatus) || CVStatus.New,
      QualificationScore: item.QualificationScore,
      ScreeningNotes: item.ScreeningNotes,
      Reviewer: item.Reviewer ? {
        Id: item.ReviewerId,
        Title: item.Reviewer.Title,
        EMail: item.Reviewer.EMail
      } : undefined,
      ReviewDate: item.ReviewDate ? new Date(item.ReviewDate) : undefined,
      IsShortlisted: item.IsShortlisted || false,
      ShortlistedBy: item.ShortlistedBy ? {
        Id: item.ShortlistedById,
        Title: item.ShortlistedBy.Title
      } : undefined,
      ShortlistedDate: item.ShortlistedDate ? new Date(item.ShortlistedDate) : undefined,
      ShortlistReason: item.ShortlistReason,
      RejectionReason: item.RejectionReason,
      RejectedBy: item.RejectedBy ? {
        Id: item.RejectedById,
        Title: item.RejectedBy.Title
      } : undefined,
      RejectionDate: item.RejectionDate ? new Date(item.RejectionDate) : undefined,
      KeywordTags: keywordTags,
      MatchScore: item.MatchScore,
      NextAction: item.NextAction,
      NextActionDate: item.NextActionDate ? new Date(item.NextActionDate) : undefined,
      InterviewScheduled: item.InterviewScheduled || false,
      InterviewDate: item.InterviewDate ? new Date(item.InterviewDate) : undefined,
      SalaryExpectation: item.SalaryExpectation,
      NoticePeriod: item.NoticePeriod,
      Availability: item.Availability,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapToDocumentRegistryEntry(item: any): IDocumentRegistryEntry {
    return {
      Id: item.Id,
      Title: item.Title,
      DocumentId: item.DocumentId || '',
      SourceModule: (item.SourceModule as SourceModule) || SourceModule.ManualUpload,
      SourceItemId: item.SourceItemId,
      SourceUrl: item.SourceUrl || item.FileUrl,
      ConfidentialityLevel: (item.ConfidentialityLevel as ConfidentialityLevel) || ConfidentialityLevel.Internal,
      DocumentStatus: (item.DocumentStatus || item.Status || DocumentStatus.Active) as DocumentStatus,
      IsRecord: item.IsRecord || false,
      OnLegalHold: item.OnLegalHold || false,
      VersionCount: item.VersionCount || 1,
      ExternalAccessEnabled: item.ExternalAccessEnabled || false,
      ActiveShareCount: item.ActiveShareCount || 0,
      ViewCount: item.ViewCount || 0,
      DownloadCount: item.DownloadCount || 0,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private async listExists(listName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Title')();
      return true;
    } catch {
      return false;
    }
  }
}

// ============================================================================
// MAIN BRIDGE SERVICE
// ============================================================================

export class DocumentHubBridgeService implements IDocumentHubModuleBridge {
  private sp: SPFI;
  private currentUserId: number = 0;
  private contractBridge: ContractManagerBridgeImpl | null = null;
  private signingBridge: SigningServiceBridgeImpl | null = null;
  private policyBridge: PolicyHubBridgeImpl | null = null;
  private cvBridge: CVManagementBridgeImpl | null = null;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize the bridge service
   */
  public async initialize(): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserId = currentUser.Id;

      // Initialize sub-bridges
      this.contractBridge = new ContractManagerBridgeImpl(this.sp, this.currentUserId);
      this.signingBridge = new SigningServiceBridgeImpl(this.sp, this.currentUserId);
      this.policyBridge = new PolicyHubBridgeImpl(this.sp, this.currentUserId);
      this.cvBridge = new CVManagementBridgeImpl(this.sp, this.currentUserId);

      logger.info('DocumentHubBridgeService', 'Bridge service initialized');
    } catch (error) {
      logger.error('DocumentHubBridgeService', 'Error initializing bridge service:', error);
      throw error;
    }
  }

  public async getAvailableModules(): Promise<IBridgeAvailability[]> {
    const availability: IBridgeAvailability[] = [];

    // Check Contract Manager
    availability.push(await this.checkModuleAvailability(
      BridgeModuleType.ContractManager,
      LISTS.CONTRACTS
    ));

    // Check Signing Service
    availability.push(await this.checkModuleAvailability(
      BridgeModuleType.SigningService,
      LISTS.DOCUMENT_SIGNING_REQUESTS
    ));

    // Check Policy Hub
    availability.push(await this.checkModuleAvailability(
      BridgeModuleType.PolicyHub,
      LISTS.POLICIES
    ));

    // Check CV Management
    availability.push(await this.checkModuleAvailability(
      BridgeModuleType.CVManagement,
      LISTS.CV_DATABASE
    ));

    return availability;
  }

  private async checkModuleAvailability(
    module: BridgeModuleType,
    listName: string
  ): Promise<IBridgeAvailability> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Title')();
      return {
        module: module,
        isAvailable: true,
        version: '1.0.0'
      };
    } catch {
      return {
        module: module,
        isAvailable: false,
        reason: `${module} module is not installed or configured`
      };
    }
  }

  public getContractManagerBridge(): IContractManagerBridge {
    if (!this.contractBridge) {
      this.contractBridge = new ContractManagerBridgeImpl(this.sp, this.currentUserId);
    }
    return this.contractBridge;
  }

  public getSigningServiceBridge(): ISigningServiceBridge {
    if (!this.signingBridge) {
      this.signingBridge = new SigningServiceBridgeImpl(this.sp, this.currentUserId);
    }
    return this.signingBridge;
  }

  public getPolicyHubBridge(): IPolicyHubBridge {
    if (!this.policyBridge) {
      this.policyBridge = new PolicyHubBridgeImpl(this.sp, this.currentUserId);
    }
    return this.policyBridge;
  }

  public getCVManagementBridge(): ICVManagementBridge {
    if (!this.cvBridge) {
      this.cvBridge = new CVManagementBridgeImpl(this.sp, this.currentUserId);
    }
    return this.cvBridge;
  }

  public async getDocumentIntegrationSummary(documentId: number): Promise<IDocumentIntegrationSummary> {
    try {
      const [contracts, signingHistory, policies, cvs, compliance] = await Promise.all([
        this.getContractManagerBridge().getLinkedContracts(documentId),
        this.getSigningServiceBridge().getSigningHistory(documentId),
        this.getPolicyHubBridge().getLinkedPolicies(documentId),
        this.getCVManagementBridge().getLinkedCVs(documentId),
        this.getPolicyHubBridge().checkPolicyCompliance(documentId)
      ]);

      const activeSigningRequest = signingHistory.some(
        r => [SigningStatus.Pending, SigningStatus.InProgress].includes(r.status)
      );

      return {
        documentId: documentId,
        linkedContracts: contracts.length,
        signingRequests: signingHistory.length,
        linkedPolicies: policies.length,
        linkedCVs: cvs.length,
        activeSigningRequest: activeSigningRequest,
        isCompliant: compliance.isCompliant,
        lastChecked: new Date()
      };
    } catch (error) {
      logger.error('DocumentHubBridgeService', 'Error getting integration summary:', error);
      return {
        documentId: documentId,
        linkedContracts: 0,
        signingRequests: 0,
        linkedPolicies: 0,
        linkedCVs: 0,
        activeSigningRequest: false,
        isCompliant: true,
        lastChecked: new Date()
      };
    }
  }

  public async getDocumentIntegrations(documentId: number): Promise<{
    contracts: IDocumentContractLink[];
    signingRequests: IDocumentSigningRequest[];
    policies: IDocumentPolicyLink[];
    cvs: IDocumentCVLink[];
    compliance: IPolicyComplianceResult | null;
  }> {
    try {
      const [contracts, signingRequests, policies, cvs, compliance] = await Promise.all([
        this.getContractManagerBridge().getLinkedContracts(documentId),
        this.getSigningServiceBridge().getSigningHistory(documentId),
        this.getPolicyHubBridge().getLinkedPolicies(documentId),
        this.getCVManagementBridge().getLinkedCVs(documentId),
        this.getPolicyHubBridge().checkPolicyCompliance(documentId)
      ]);

      return {
        contracts,
        signingRequests,
        policies,
        cvs,
        compliance
      };
    } catch (error) {
      logger.error('DocumentHubBridgeService', 'Error getting document integrations:', error);
      return {
        contracts: [],
        signingRequests: [],
        policies: [],
        cvs: [],
        compliance: null
      };
    }
  }
}
