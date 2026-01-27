// @ts-nocheck
// E-Signature Integration Service
// Handles integration with DocuSign and Adobe Sign for document signatures

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  ISignatureRequest,
  ISignatureRecipient,
  ISignatureWebhookEvent,
  SignatureProvider,
  SignatureStatus
} from '../models';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class ESignatureService {
  private sp: SPFI;
  private readonly SIGNATURE_CONFIG_LIST = 'JML_SignatureConfig';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Send document for signature
   */
  public async sendForSignature(request: ISignatureRequest): Promise<string> {
    try {
      switch (request.provider) {
        case SignatureProvider.DocuSign:
          return await this.sendDocuSignEnvelope(request);
        case SignatureProvider.AdobeSign:
          return await this.sendAdobeSignAgreement(request);
        case SignatureProvider.Internal:
          return await this.sendInternalSignatureRequest(request);
        default:
          throw new Error(`Unsupported signature provider: ${request.provider}`);
      }
    } catch (error) {
      logger.error('ESignatureService', 'Failed to send for signature:', error);
      throw error;
    }
  }

  /**
   * Send DocuSign envelope
   */
  private async sendDocuSignEnvelope(request: ISignatureRequest): Promise<string> {
    try {
      const config = await this.getSignatureConfig(SignatureProvider.DocuSign);

      // Prepare envelope definition
      const envelopeDefinition = {
        emailSubject: request.emailSubject,
        emailMessage: request.emailMessage,
        recipients: {
          signers: this.mapRecipientsToDocuSign(request.signers)
        },
        documents: [{
          documentId: request.documentId.toString(),
          name: `Document_${request.documentId}`,
          documentBase64: await this.getDocumentBase64(request.documentId)
        }],
        status: 'sent',
        notification: {
          useAccountDefaults: false,
          reminders: {
            reminderEnabled: request.reminderDays ? true : false,
            reminderDelay: request.reminderDays?.toString() || '3',
            reminderFrequency: '1'
          },
          expirations: {
            expireEnabled: request.expirationDays ? true : false,
            expireAfter: request.expirationDays?.toString() || '30',
            expireWarn: '3'
          }
        }
      };

      // Call DocuSign API
      const response = await fetch(`${config.apiBaseUrl}/v2.1/accounts/${config.accountId}/envelopes`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${config.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(envelopeDefinition)
      });

      if (!response.ok) {
        throw new Error(`DocuSign API error: ${response.statusText}`);
      }

      const result = await response.json();

      // Update document with envelope ID
      await this.updateDocumentSignatureInfo(
        request.documentId,
        result.envelopeId,
        SignatureStatus.Sent,
        SignatureProvider.DocuSign
      );

      return result.envelopeId;
    } catch (error) {
      logger.error('ESignatureService', 'Failed to send DocuSign envelope:', error);
      throw error;
    }
  }

  /**
   * Send Adobe Sign agreement
   */
  private async sendAdobeSignAgreement(request: ISignatureRequest): Promise<string> {
    try {
      const config = await this.getSignatureConfig(SignatureProvider.AdobeSign);

      // Prepare transient document
      const documentBase64 = await this.getDocumentBase64(request.documentId);
      const transientDocumentId = await this.uploadTransientDocument(
        config,
        documentBase64,
        `Document_${request.documentId}.pdf`
      );

      // Prepare agreement info
      const agreementInfo = {
        fileInfos: [{
          transientDocumentId: transientDocumentId
        }],
        name: request.emailSubject,
        participantSetsInfo: this.mapRecipientsToAdobeSign(request.signers),
        signatureType: 'ESIGN',
        state: 'IN_PROCESS',
        emailOption: {
          sendOptions: {
            initEmails: 'ALL'
          }
        },
        message: request.emailMessage,
        reminderFrequency: request.reminderDays ? 'DAILY_UNTIL_SIGNED' : 'NONE',
        externalId: {
          id: request.documentId.toString()
        }
      };

      // Call Adobe Sign API
      const response = await fetch(`${config.apiBaseUrl}/api/rest/v6/agreements`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${config.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(agreementInfo)
      });

      if (!response.ok) {
        throw new Error(`Adobe Sign API error: ${response.statusText}`);
      }

      const result = await response.json();

      // Update document with agreement ID
      await this.updateDocumentSignatureInfo(
        request.documentId,
        result.id,
        SignatureStatus.Sent,
        SignatureProvider.AdobeSign
      );

      return result.id;
    } catch (error) {
      logger.error('ESignatureService', 'Failed to send Adobe Sign agreement:', error);
      throw error;
    }
  }

  /**
   * Send internal signature request
   */
  private async sendInternalSignatureRequest(request: ISignatureRequest): Promise<string> {
    try {
      // Create signature request record
      const signatureRequest = await this.sp.web.lists
        .getByTitle('JML_SignatureRequests')
        .items.add({
          DocumentId: request.documentId,
          ProcessId: request.processId,
          EmailSubject: request.emailSubject,
          EmailMessage: request.emailMessage,
          Recipients: JSON.stringify(request.signers),
          Status: SignatureStatus.Sent,
          ExpirationDate: request.expirationDays
            ? new Date(Date.now() + request.expirationDays * 24 * 60 * 60 * 1000).toISOString()
            : undefined
        });

      // Update document
      await this.updateDocumentSignatureInfo(
        request.documentId,
        signatureRequest.data.Id.toString(),
        SignatureStatus.Sent,
        SignatureProvider.Internal
      );

      // Send email notifications
      await this.sendSignatureRequestEmails(request);

      return signatureRequest.data.Id.toString();
    } catch (error) {
      logger.error('ESignatureService', 'Failed to send internal signature request:', error);
      throw error;
    }
  }

  /**
   * Get signature status
   */
  public async getSignatureStatus(envelopeId: string, provider: SignatureProvider): Promise<SignatureStatus> {
    try {
      switch (provider) {
        case SignatureProvider.DocuSign:
          return await this.getDocuSignStatus(envelopeId);
        case SignatureProvider.AdobeSign:
          return await this.getAdobeSignStatus(envelopeId);
        case SignatureProvider.Internal:
          return await this.getInternalSignatureStatus(parseInt(envelopeId, 10));
        default:
          return SignatureStatus.NotRequired;
      }
    } catch (error) {
      logger.error('ESignatureService', 'Failed to get signature status:', error);
      return SignatureStatus.Pending;
    }
  }

  /**
   * Process signature webhook event
   */
  public async processWebhookEvent(event: ISignatureWebhookEvent): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_Documents')
        .items.getById(event.documentId)
        .update({
          SignatureStatus: event.status,
          SignedDate: event.signedDate?.toISOString(),
          SignedBy: event.signedBy
        });
    } catch (error) {
      logger.error('ESignatureService', 'Failed to process webhook event:', error);
      throw error;
    }
  }

  /**
   * Void/cancel signature request
   */
  public async voidSignatureRequest(envelopeId: string, provider: SignatureProvider, reason: string): Promise<void> {
    try {
      switch (provider) {
        case SignatureProvider.DocuSign:
          await this.voidDocuSignEnvelope(envelopeId, reason);
          break;
        case SignatureProvider.AdobeSign:
          await this.voidAdobeSignAgreement(envelopeId, reason);
          break;
        case SignatureProvider.Internal:
          await this.voidInternalSignatureRequest(parseInt(envelopeId, 10), reason);
          break;
      }
    } catch (error) {
      logger.error('ESignatureService', 'Failed to void signature request:', error);
      throw error;
    }
  }

  // Private helper methods

  private async getSignatureConfig(provider: SignatureProvider): Promise<any> {
    // Validate enum value
    ValidationUtils.validateEnum(provider, SignatureProvider, 'SignatureProvider');

    // Build secure filter
    const filter = ValidationUtils.buildFilter('Provider', 'eq', provider);

    const items = await this.sp.web.lists
      .getByTitle(this.SIGNATURE_CONFIG_LIST)
      .items.filter(filter)
      .top(1)();

    if (items.length === 0) {
      throw new Error(`No configuration found for provider: ${provider}`);
    }

    return {
      provider,
      apiBaseUrl: items[0].ApiBaseUrl,
      accountId: items[0].AccountId,
      accessToken: items[0].AccessToken,
      webhookUrl: items[0].WebhookUrl
    };
  }

  private async getDocumentBase64(documentId: number): Promise<string> {
    const item = await this.sp.web.lists
      .getByTitle('JML_Documents')
      .items.getById(documentId)
      .select('File/ServerRelativeUrl')
      .expand('File')();

    const fileUrl = `${window.location.origin}${item.File.ServerRelativeUrl}`;
    const response = await fetch(fileUrl);
    const blob = await response.blob();

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        resolve(base64.split(',')[1]);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  private mapRecipientsToDocuSign(signers: ISignatureRecipient[]): any[] {
    const docuSignSigners: any[] = [];

    for (let i = 0; i < signers.length; i++) {
      const signer = signers[i];
      if (signer.role === 'Signer') {
        docuSignSigners.push({
          email: signer.email,
          name: signer.name,
          recipientId: (i + 1).toString(),
          routingOrder: signer.routingOrder.toString(),
          idCheckConfigurationName: signer.requireIdVerification ? 'ID Check $' : undefined
        });
      }
    }

    return docuSignSigners;
  }

  private mapRecipientsToAdobeSign(signers: ISignatureRecipient[]): any[] {
    const participantSets: any[] = [];

    for (let i = 0; i < signers.length; i++) {
      const signer = signers[i];
      if (signer.role === 'Signer') {
        participantSets.push({
          memberInfos: [{
            email: signer.email,
            name: signer.name
          }],
          order: signer.routingOrder,
          role: 'SIGNER'
        });
      }
    }

    return participantSets;
  }

  private async uploadTransientDocument(config: any, base64Content: string, fileName: string): Promise<string> {
    const formData = new FormData();
    const blob = this.base64ToBlob(base64Content, 'application/pdf');
    formData.append('File', blob, fileName);

    const response = await fetch(`${config.apiBaseUrl}/api/rest/v6/transientDocuments`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`
      },
      body: formData
    });

    if (!response.ok) {
      throw new Error('Failed to upload transient document');
    }

    const result = await response.json();
    return result.transientDocumentId;
  }

  private base64ToBlob(base64: string, contentType: string): Blob {
    const byteCharacters = atob(base64);
    const byteNumbers: number[] = [];

    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers.push(byteCharacters.charCodeAt(i));
    }

    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: contentType });
  }

  private async updateDocumentSignatureInfo(
    documentId: number,
    envelopeId: string,
    status: SignatureStatus,
    provider: SignatureProvider
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle('JML_Documents')
      .items.getById(documentId)
      .update({
        SignatureEnvelopeId: envelopeId,
        SignatureStatus: status,
        SignatureProvider: provider
      });
  }

  private async getDocuSignStatus(envelopeId: string): Promise<SignatureStatus> {
    const config = await this.getSignatureConfig(SignatureProvider.DocuSign);

    const response = await fetch(
      `${config.apiBaseUrl}/v2.1/accounts/${config.accountId}/envelopes/${envelopeId}`,
      {
        headers: {
          'Authorization': `Bearer ${config.accessToken}`
        }
      }
    );

    if (!response.ok) {
      throw new Error('Failed to get envelope status');
    }

    const envelope = await response.json();
    return this.mapDocuSignStatus(envelope.status);
  }

  private async getAdobeSignStatus(agreementId: string): Promise<SignatureStatus> {
    const config = await this.getSignatureConfig(SignatureProvider.AdobeSign);

    const response = await fetch(
      `${config.apiBaseUrl}/api/rest/v6/agreements/${agreementId}`,
      {
        headers: {
          'Authorization': `Bearer ${config.accessToken}`
        }
      }
    );

    if (!response.ok) {
      throw new Error('Failed to get agreement status');
    }

    const agreement = await response.json();
    return this.mapAdobeSignStatus(agreement.status);
  }

  private async getInternalSignatureStatus(requestId: number): Promise<SignatureStatus> {
    const item = await this.sp.web.lists
      .getByTitle('JML_SignatureRequests')
      .items.getById(requestId)
      .select('Status')();

    return item.Status as SignatureStatus;
  }

  private async voidDocuSignEnvelope(envelopeId: string, reason: string): Promise<void> {
    const config = await this.getSignatureConfig(SignatureProvider.DocuSign);

    await fetch(
      `${config.apiBaseUrl}/v2.1/accounts/${config.accountId}/envelopes/${envelopeId}`,
      {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${config.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          status: 'voided',
          voidedReason: reason
        })
      }
    );
  }

  private async voidAdobeSignAgreement(agreementId: string, reason: string): Promise<void> {
    const config = await this.getSignatureConfig(SignatureProvider.AdobeSign);

    await fetch(
      `${config.apiBaseUrl}/api/rest/v6/agreements/${agreementId}/state`,
      {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${config.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          state: 'CANCELLED',
          agreementCancellationInfo: {
            comment: reason
          }
        })
      }
    );
  }

  private async voidInternalSignatureRequest(requestId: number, reason: string): Promise<void> {
    await this.sp.web.lists
      .getByTitle('JML_SignatureRequests')
      .items.getById(requestId)
      .update({
        Status: SignatureStatus.Voided,
        VoidReason: reason
      });
  }

  private mapDocuSignStatus(status: string): SignatureStatus {
    switch (status.toLowerCase()) {
      case 'completed':
        return SignatureStatus.Signed;
      case 'sent':
      case 'delivered':
        return SignatureStatus.Sent;
      case 'declined':
        return SignatureStatus.Declined;
      case 'voided':
        return SignatureStatus.Voided;
      default:
        return SignatureStatus.Pending;
    }
  }

  private mapAdobeSignStatus(status: string): SignatureStatus {
    switch (status) {
      case 'SIGNED':
        return SignatureStatus.Signed;
      case 'OUT_FOR_SIGNATURE':
        return SignatureStatus.Sent;
      case 'RECALLED':
      case 'CANCELLED':
        return SignatureStatus.Voided;
      case 'EXPIRED':
        return SignatureStatus.Expired;
      default:
        return SignatureStatus.Pending;
    }
  }

  private async sendSignatureRequestEmails(request: ISignatureRequest): Promise<void> {
    // This would integrate with your email service
    // For now, this is a placeholder
    logger.debug('ESignatureService', 'Sending signature request emails to:', { data: request.signers });
  }
}
