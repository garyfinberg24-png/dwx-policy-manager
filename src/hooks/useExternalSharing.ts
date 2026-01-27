// @ts-nocheck
/**
 * useExternalSharing Hook
 *
 * Provides External Sharing Hub functionality for integration with other JML modules
 * such as Contract Manager, Procurement Manager, Signing Service, and Policy Hub.
 *
 * Usage in Contract Manager:
 * ```tsx
 * const { shareResource, inviteGuest, trustedOrganizations } = useExternalSharing(context);
 *
 * // Share a contract with an external party
 * await shareResource({
 *   title: 'Contract #12345',
 *   resourceUrl: contract.documentUrl,
 *   resourceType: ResourceType.Document,
 *   relatedModule: 'Contract',
 *   relatedItemId: contract.Id.toString()
 * });
 * ```
 */

import { useState, useEffect, useCallback, useMemo } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as SPFxSP } from '@pnp/sp';
import { graphfi, SPFx as SPFxGraph } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/graph/users';
import { ExternalSharingService } from '../services/ExternalSharingService';
import {
  ITrustedOrganization,
  ISharedResource,
  IShareRequest,
  IInvitation,
  IInvitationResult,
  ResourceType,
  SharingLevel,
  DataClassification,
  TrustStatus,
  RelatedModule
} from '../models/IExternalSharing';

export interface IUseExternalSharingOptions {
  /** Auto-load trusted organizations on mount */
  autoLoadOrganizations?: boolean;
  /** Filter organizations by status */
  organizationStatusFilter?: TrustStatus;
}

export interface IUseExternalSharingResult {
  /** Service instance for advanced operations */
  service: ExternalSharingService;

  /** List of trusted organizations */
  trustedOrganizations: ITrustedOrganization[];

  /** Loading state for organizations */
  isLoadingOrganizations: boolean;

  /** Error state */
  error: string | null;

  /** Refresh trusted organizations */
  refreshOrganizations: () => Promise<void>;

  /**
   * Share a resource with an external organization
   */
  shareResource: (request: IShareResourceRequest) => Promise<ISharedResource | null>;

  /**
   * Invite a guest user from a trusted organization
   */
  inviteGuest: (invitation: IInviteGuestRequest) => Promise<IInvitationResult>;

  /**
   * Check if a domain is from a trusted organization
   */
  isDomainTrusted: (domain: string) => boolean;

  /**
   * Get trusted organization by domain
   */
  getOrganizationByDomain: (domain: string) => ITrustedOrganization | undefined;

  /**
   * Validate if sharing is allowed based on classification
   */
  validateSharingAllowed: (
    organizationId: number,
    classification: DataClassification,
    sharingLevel: SharingLevel
  ) => { allowed: boolean; reason?: string };
}

/**
 * Simplified share request for module integration
 */
export interface IShareResourceRequest {
  /** Resource title */
  title: string;
  /** URL to the resource */
  resourceUrl: string;
  /** Type of resource */
  resourceType: ResourceType;
  /** Organization ID to share with */
  organizationId?: number;
  /** Specific user emails to share with */
  userEmails?: string[];
  /** Permission level */
  sharingLevel?: SharingLevel;
  /** Data classification */
  dataClassification?: DataClassification;
  /** Expiration in days (default: from org policy) */
  expirationDays?: number;
  /** Require acknowledgment before access */
  requiresAcknowledgment?: boolean;
  /** Related JML module */
  relatedModule: 'Contract' | 'Procurement' | 'Signing' | 'Policy' | 'Other';
  /** ID of related item in the module */
  relatedItemId?: string;
}

/**
 * Simplified guest invitation for module integration
 */
export interface IInviteGuestRequest {
  /** Guest email address */
  email: string;
  /** Display name */
  displayName: string;
  /** Organization ID (will be validated against trusted orgs) */
  organizationId: number;
  /** Send invitation email */
  sendInvitation?: boolean;
  /** Custom invitation message */
  customMessage?: string;
}

/**
 * Hook for integrating External Sharing Hub functionality into other modules
 */
export function useExternalSharing(
  context: WebPartContext,
  options: IUseExternalSharingOptions = {}
): IUseExternalSharingResult {
  const {
    autoLoadOrganizations = true,
    organizationStatusFilter = TrustStatus.Active
  } = options;

  const [trustedOrganizations, setTrustedOrganizations] = useState<ITrustedOrganization[]>([]);
  const [isLoadingOrganizations, setIsLoadingOrganizations] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  // Initialize PnP instances and service
  const service = useMemo(() => {
    const sp = spfi().using(SPFxSP(context));
    const graph = graphfi().using(SPFxGraph(context));
    return new ExternalSharingService(sp, graph);
  }, [context]);

  // Load trusted organizations
  const refreshOrganizations = useCallback(async (): Promise<void> => {
    try {
      setIsLoadingOrganizations(true);
      setError(null);

      const orgs = await service.getTrustedOrganizations();

      // Filter by status if specified
      const filteredOrgs = organizationStatusFilter
        ? orgs.filter(org => org.Status === organizationStatusFilter)
        : orgs;

      setTrustedOrganizations(filteredOrgs);
    } catch (err) {
      console.error('Failed to load trusted organizations:', err);
      setError(err instanceof Error ? err.message : 'Failed to load organizations');
    } finally {
      setIsLoadingOrganizations(false);
    }
  }, [service, organizationStatusFilter]);

  // Auto-load on mount
  useEffect(() => {
    if (autoLoadOrganizations) {
      refreshOrganizations().catch(console.error);
    }
  }, [autoLoadOrganizations, refreshOrganizations]);

  // Check if domain is trusted
  const isDomainTrusted = useCallback((domain: string): boolean => {
    const normalizedDomain = domain.toLowerCase();
    return trustedOrganizations.some(org =>
      org.TenantDomain.toLowerCase() === normalizedDomain ||
      (org.AllowedDomains && typeof org.AllowedDomains === 'string' &&
        JSON.parse(org.AllowedDomains).some((d: string) => d.toLowerCase() === normalizedDomain))
    );
  }, [trustedOrganizations]);

  // Get organization by domain
  const getOrganizationByDomain = useCallback((domain: string): ITrustedOrganization | undefined => {
    const normalizedDomain = domain.toLowerCase();
    return trustedOrganizations.find(org =>
      org.TenantDomain.toLowerCase() === normalizedDomain ||
      (org.AllowedDomains && typeof org.AllowedDomains === 'string' &&
        JSON.parse(org.AllowedDomains).some((d: string) => d.toLowerCase() === normalizedDomain))
    );
  }, [trustedOrganizations]);

  // Validate sharing allowed
  const validateSharingAllowed = useCallback((
    organizationId: number,
    classification: DataClassification,
    sharingLevel: SharingLevel
  ): { allowed: boolean; reason?: string } => {
    const org = trustedOrganizations.find(o => o.Id === organizationId);

    if (!org) {
      return { allowed: false, reason: 'Organization not found or not trusted' };
    }

    if (org.Status !== TrustStatus.Active) {
      return { allowed: false, reason: `Organization trust status is ${org.Status}` };
    }

    // Check classification restrictions
    if (classification === DataClassification.HighlyConfidential) {
      return { allowed: false, reason: 'Highly Confidential data requires executive approval' };
    }

    if (classification === DataClassification.Confidential && org.TrustLevel !== 'Full') {
      return { allowed: false, reason: 'Confidential data requires Full Trust relationship' };
    }

    // Check sharing level restrictions
    const sharingLevels = [SharingLevel.View, SharingLevel.Edit, SharingLevel.FullControl];
    const requestedLevel = sharingLevels.indexOf(sharingLevel);
    const maxLevel = sharingLevels.indexOf(org.MaxSharingLevel);

    if (requestedLevel > maxLevel) {
      return {
        allowed: false,
        reason: `Maximum sharing level for ${org.Title} is ${org.MaxSharingLevel}`
      };
    }

    return { allowed: true };
  }, [trustedOrganizations]);

  // Share resource
  const shareResource = useCallback(async (
    request: IShareResourceRequest
  ): Promise<ISharedResource | null> => {
    try {
      setError(null);

      // Validate if organization is specified
      if (request.organizationId) {
        const validation = validateSharingAllowed(
          request.organizationId,
          request.dataClassification || DataClassification.Internal,
          request.sharingLevel || SharingLevel.View
        );

        if (!validation.allowed) {
          throw new Error(validation.reason);
        }
      }

      const resourceService = service.getSharedResourceService();

      const shareRequest: IShareRequest = {
        title: request.title,
        resourceUrl: request.resourceUrl,
        resourceType: request.resourceType,
        sharedWithOrganizationId: request.organizationId,
        sharedWithUsers: request.userEmails,
        sharingLevel: request.sharingLevel || SharingLevel.View,
        dataClassification: request.dataClassification || DataClassification.Internal,
        requiresAcknowledgment: request.requiresAcknowledgment || false,
        relatedModule: request.relatedModule,
        relatedItemId: request.relatedItemId,
        expirationDays: request.expirationDays
      };

      const result = await resourceService.shareResource(shareRequest);
      return result;
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to share resource';
      setError(errorMessage);
      console.error('Share resource error:', err);
      return null;
    }
  }, [service, validateSharingAllowed]);

  // Invite guest
  const inviteGuest = useCallback(async (
    request: IInviteGuestRequest
  ): Promise<IInvitationResult> => {
    try {
      setError(null);

      // Validate organization is trusted
      const org = trustedOrganizations.find(o => o.Id === request.organizationId);
      if (!org || org.Status !== TrustStatus.Active) {
        return {
          success: false,
          error: 'Organization is not trusted or not active'
        };
      }

      // Validate email domain
      const emailDomain = request.email.split('@')[1]?.toLowerCase();
      if (emailDomain !== org.TenantDomain.toLowerCase()) {
        // Check allowed domains
        let domainAllowed = false;
        if (org.AllowedDomains && typeof org.AllowedDomains === 'string') {
          const allowedDomains = JSON.parse(org.AllowedDomains);
          domainAllowed = allowedDomains.some((d: string) => d.toLowerCase() === emailDomain);
        }

        if (!domainAllowed) {
          return {
            success: false,
            error: `Email domain ${emailDomain} is not allowed for organization ${org.Title}`
          };
        }
      }

      const guestService = service.getGuestUserService();

      const invitation: IInvitation = {
        email: request.email,
        displayName: request.displayName,
        sourceOrganizationId: request.organizationId,
        sendInvitationMessage: request.sendInvitation !== false,
        invitationMessage: request.customMessage
      };

      const result = await guestService.inviteGuestUser(invitation);
      return result;
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to invite guest';
      setError(errorMessage);
      console.error('Invite guest error:', err);
      return {
        success: false,
        error: errorMessage
      };
    }
  }, [service, trustedOrganizations]);

  return {
    service,
    trustedOrganizations,
    isLoadingOrganizations,
    error,
    refreshOrganizations,
    shareResource,
    inviteGuest,
    isDomainTrusted,
    getOrganizationByDomain,
    validateSharingAllowed
  };
}

// Re-export types for convenience
export {
  ResourceType,
  SharingLevel,
  DataClassification,
  TrustStatus,
  RelatedModule
} from '../models/IExternalSharing';
