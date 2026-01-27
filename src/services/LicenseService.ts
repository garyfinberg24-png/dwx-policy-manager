// @ts-nocheck
/**
 * LicenseService - Premium Module Licensing
 *
 * Manages license validation, module access control, and usage tracking.
 *
 * SECURITY NOTES:
 * - Client-side validation provides UI-level protection
 * - License data is cached locally but validated against SharePoint
 * - Usage is logged for audit purposes
 * - For enhanced security, implement server-side validation via Azure Functions
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  ILicense,
  ILicenseData,
  IModuleLicenseCheck,
  ILicenseActivationRequest,
  ILicenseActivationResponse,
  PremiumModule,
  LicenseTier,
  LicenseStatus,
  TierModules
} from '../models/ILicense';
import { logger } from './LoggingService';

// Cache duration in milliseconds (15 minutes)
const CACHE_DURATION = 15 * 60 * 1000;

// Local storage keys
const LICENSE_CACHE_KEY = 'jml_license_cache';
const LICENSE_CACHE_TIMESTAMP_KEY = 'jml_license_cache_ts';

export class LicenseService {
  private sp: SPFI;
  private cachedLicense: ILicenseData | null = null;
  private cacheTimestamp: number = 0;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.loadFromLocalStorage();
  }

  // ===================================================================
  // License Retrieval
  // ===================================================================

  /**
   * Get the current license data for this tenant
   * @param forceRefresh - Bypass cache and fetch fresh data
   */
  public async getLicense(forceRefresh: boolean = false): Promise<ILicenseData> {
    try {
      // Check cache first
      if (!forceRefresh && this.isCacheValid()) {
        logger.info('LicenseService', 'Returning cached license data');
        return this.cachedLicense!;
      }

      // Fetch from SharePoint
      const license = await this.fetchLicenseFromSharePoint();

      if (license) {
        const licenseData = this.parseLicense(license);
        this.cacheLocalLicense(licenseData);

        // Log license check for audit
        await this.logLicenseCheck(license.LicenseKey, 'LICENSE_VALIDATED');

        return licenseData;
      }

      // No license found - return free tier
      const freeLicense = this.getFreeTierLicense();
      this.cacheLocalLicense(freeLicense);
      return freeLicense;

    } catch (error) {
      logger.error('LicenseService', 'Error fetching license', error);

      // On error, return cached data if available, otherwise free tier
      if (this.cachedLicense) {
        return this.cachedLicense;
      }
      return this.getFreeTierLicense();
    }
  }

  /**
   * Check if a specific module is licensed
   */
  public async isModuleLicensed(moduleId: PremiumModule): Promise<IModuleLicenseCheck> {
    const license = await this.getLicense();

    if (!license.isValid) {
      return {
        moduleId,
        isLicensed: false,
        reason: 'no_license'
      };
    }

    if (license.status === LicenseStatus.Expired) {
      return {
        moduleId,
        isLicensed: false,
        reason: 'license_expired'
      };
    }

    if (license.status === LicenseStatus.Suspended) {
      return {
        moduleId,
        isLicensed: false,
        reason: 'license_suspended'
      };
    }

    if (license.status === LicenseStatus.Trial && license.daysUntilExpiration !== undefined && license.daysUntilExpiration <= 0) {
      return {
        moduleId,
        isLicensed: false,
        reason: 'trial_ended'
      };
    }

    const isEnabled = license.enabledModules.includes(moduleId);

    return {
      moduleId,
      isLicensed: isEnabled,
      reason: isEnabled ? undefined : 'not_in_tier'
    };
  }

  /**
   * Check multiple modules at once
   */
  public async checkModules(moduleIds: PremiumModule[]): Promise<IModuleLicenseCheck[]> {
    const license = await this.getLicense();
    return moduleIds.map(moduleId => {
      const isLicensed = license.isValid &&
        license.status === LicenseStatus.Active &&
        license.enabledModules.includes(moduleId);

      return {
        moduleId,
        isLicensed,
        reason: isLicensed ? undefined : 'not_in_tier'
      };
    });
  }

  /**
   * Get all licensed modules
   */
  public async getLicensedModules(): Promise<PremiumModule[]> {
    const license = await this.getLicense();
    return license.enabledModules;
  }

  // ===================================================================
  // License Activation
  // ===================================================================

  /**
   * Activate a license key
   */
  public async activateLicense(request: ILicenseActivationRequest): Promise<ILicenseActivationResponse> {
    try {
      // Validate license key format
      if (!this.isValidLicenseKeyFormat(request.licenseKey)) {
        return {
          success: false,
          message: 'Invalid license key format',
          errorCode: 'INVALID_KEY'
        };
      }

      // Check if license already exists for this tenant
      const existingLicense = await this.fetchLicenseFromSharePoint();
      if (existingLicense) {
        // Update existing license
        return await this.updateExistingLicense(existingLicense, request);
      }

      // Create new license entry
      const newLicense: Partial<ILicense> = {
        Title: `License - ${request.organizationName || request.tenantId}`,
        LicenseKey: request.licenseKey,
        TenantId: request.tenantId,
        OrganizationName: request.organizationName || 'Unknown',
        ContactEmail: request.contactEmail,
        Status: LicenseStatus.PendingActivation,
        Tier: LicenseTier.Free, // Will be updated by admin
        EnabledModules: '[]',
        MaxUsers: 0,
        ActivatedDate: new Date()
      };

      await this.sp.web.lists.getByTitle('JML_Licenses').items.add(newLicense);

      // Log activation attempt
      await this.logLicenseCheck(request.licenseKey, 'ACTIVATION_REQUESTED');

      // Clear cache to force refresh
      this.clearCache();

      return {
        success: true,
        message: 'License activation requested. Please contact support to complete activation.',
        license: this.getFreeTierLicense()
      };

    } catch (error) {
      logger.error('LicenseService', 'Error activating license', error);
      return {
        success: false,
        message: 'An error occurred during license activation',
        errorCode: 'SERVER_ERROR'
      };
    }
  }

  // ===================================================================
  // Usage Tracking (for audit)
  // ===================================================================

  /**
   * Log premium module usage
   */
  public async logModuleUsage(moduleId: PremiumModule, action: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_AuditLog').items.add({
        Title: `Module Usage: ${moduleId}`,
        Action: action,
        Category: 'LICENSE',
        Details: JSON.stringify({
          moduleId,
          action,
          timestamp: new Date().toISOString()
        }),
        Timestamp: new Date()
      });
    } catch (error) {
      // Don't fail on logging errors - just log locally
      logger.warn('LicenseService', 'Failed to log module usage', error);
    }
  }

  // ===================================================================
  // Admin Functions
  // ===================================================================

  /**
   * Get all licenses (admin only)
   */
  public async getAllLicenses(): Promise<ILicense[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('JML_Licenses').items
        .select('*')
        .orderBy('Created', false)();

      return items as ILicense[];
    } catch (error) {
      logger.error('LicenseService', 'Error fetching all licenses', error);
      throw error;
    }
  }

  /**
   * Update a license (admin only)
   */
  public async updateLicense(id: number, updates: Partial<ILicense>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_Licenses').items.getById(id).update(updates);

      // Clear cache after update
      this.clearCache();

      // Log the update
      await this.logLicenseCheck(updates.LicenseKey || 'unknown', 'LICENSE_UPDATED');

    } catch (error) {
      logger.error('LicenseService', 'Error updating license', error);
      throw error;
    }
  }

  /**
   * Generate a new license key (admin only)
   * Format: JML-{TIER}-{YEAR}-{RANDOM}
   */
  public generateLicenseKey(tier: LicenseTier): string {
    const tierCode = tier.substring(0, 3).toUpperCase();
    const year = new Date().getFullYear();
    const random = this.generateRandomString(8);
    return `JML-${tierCode}-${year}-${random}`;
  }

  // ===================================================================
  // Private Methods
  // ===================================================================

  private async fetchLicenseFromSharePoint(): Promise<ILicense | null> {
    try {
      const items = await this.sp.web.lists.getByTitle('JML_Licenses').items
        .select('*')
        .filter(`Status ne '${LicenseStatus.Suspended}'`)
        .top(1)();

      if (items.length > 0) {
        return items[0] as ILicense;
      }
      return null;
    } catch (error) {
      // List might not exist yet
      logger.warn('LicenseService', 'JML_Licenses list not found or error', error);
      return null;
    }
  }

  private parseLicense(license: ILicense): ILicenseData {
    const now = new Date();
    const expirationDate = license.ExpirationDate ? new Date(license.ExpirationDate) : undefined;
    const isExpired = expirationDate ? expirationDate < now : false;
    const daysUntilExpiration = expirationDate
      ? Math.ceil((expirationDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24))
      : undefined;

    // Parse enabled modules
    let enabledModules: PremiumModule[] = [];
    try {
      if (license.Tier === LicenseTier.Custom) {
        enabledModules = JSON.parse(license.EnabledModules || '[]');
      } else {
        enabledModules = TierModules[license.Tier] || [];
      }
    } catch {
      enabledModules = TierModules[license.Tier] || [];
    }

    return {
      isValid: license.Status === LicenseStatus.Active || license.Status === LicenseStatus.Trial,
      tier: license.Tier,
      status: isExpired ? LicenseStatus.Expired : license.Status,
      enabledModules,
      expirationDate,
      daysUntilExpiration,
      maxUsers: license.MaxUsers,
      organizationName: license.OrganizationName,
      isTrial: license.Status === LicenseStatus.Trial,
      isExpiringSoon: daysUntilExpiration !== undefined && daysUntilExpiration <= 30 && daysUntilExpiration > 0
    };
  }

  private getFreeTierLicense(): ILicenseData {
    return {
      isValid: true,
      tier: LicenseTier.Free,
      status: LicenseStatus.Active,
      enabledModules: [],
      maxUsers: 0,
      organizationName: '',
      isTrial: false,
      isExpiringSoon: false
    };
  }

  private async updateExistingLicense(existing: ILicense, request: ILicenseActivationRequest): Promise<ILicenseActivationResponse> {
    // If key matches, just update contact info
    if (existing.LicenseKey === request.licenseKey) {
      await this.sp.web.lists.getByTitle('JML_Licenses').items.getById(existing.Id!).update({
        ContactEmail: request.contactEmail,
        LastValidated: new Date()
      });

      this.clearCache();

      return {
        success: true,
        message: 'License validated successfully',
        license: this.parseLicense(existing)
      };
    }

    // Different key - update to new key (pending admin approval)
    await this.sp.web.lists.getByTitle('JML_Licenses').items.getById(existing.Id!).update({
      LicenseKey: request.licenseKey,
      ContactEmail: request.contactEmail,
      Status: LicenseStatus.PendingActivation,
      Notes: `Key changed from ${existing.LicenseKey} on ${new Date().toISOString()}`
    });

    this.clearCache();

    return {
      success: true,
      message: 'New license key submitted for activation'
    };
  }

  private async logLicenseCheck(licenseKey: string, action: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_AuditLog').items.add({
        Title: `License Check: ${action}`,
        Action: action,
        Category: 'LICENSE',
        Details: JSON.stringify({
          licenseKey: this.maskLicenseKey(licenseKey),
          action,
          timestamp: new Date().toISOString()
        }),
        Timestamp: new Date()
      });
    } catch {
      // Silent fail for logging
    }
  }

  private maskLicenseKey(key: string): string {
    if (!key || key.length < 8) return '****';
    return `${key.substring(0, 4)}...${key.substring(key.length - 4)}`;
  }

  private isValidLicenseKeyFormat(key: string): boolean {
    // Format: JML-XXX-YYYY-XXXXXXXX
    const pattern = /^JML-[A-Z]{3}-\d{4}-[A-Z0-9]{8}$/;
    return pattern.test(key);
  }

  private generateRandomString(length: number): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }

  // ===================================================================
  // Caching
  // ===================================================================

  private isCacheValid(): boolean {
    if (!this.cachedLicense) return false;
    const now = Date.now();
    return (now - this.cacheTimestamp) < CACHE_DURATION;
  }

  private cacheLocalLicense(license: ILicenseData): void {
    this.cachedLicense = license;
    this.cacheTimestamp = Date.now();

    // Also store in localStorage for persistence
    try {
      localStorage.setItem(LICENSE_CACHE_KEY, JSON.stringify(license));
      localStorage.setItem(LICENSE_CACHE_TIMESTAMP_KEY, this.cacheTimestamp.toString());
    } catch {
      // localStorage might be unavailable
    }
  }

  private loadFromLocalStorage(): void {
    try {
      const cached = localStorage.getItem(LICENSE_CACHE_KEY);
      const timestamp = localStorage.getItem(LICENSE_CACHE_TIMESTAMP_KEY);

      if (cached && timestamp) {
        this.cachedLicense = JSON.parse(cached);
        this.cacheTimestamp = parseInt(timestamp, 10);
      }
    } catch {
      // Ignore localStorage errors
    }
  }

  private clearCache(): void {
    this.cachedLicense = null;
    this.cacheTimestamp = 0;
    try {
      localStorage.removeItem(LICENSE_CACHE_KEY);
      localStorage.removeItem(LICENSE_CACHE_TIMESTAMP_KEY);
    } catch {
      // Ignore
    }
  }
}
