/**
 * useLicense - React hook for premium module licensing
 *
 * Provides easy access to license status and module checks throughout the app.
 */

import { useState, useEffect, useCallback, useMemo } from 'react';
import { LicenseService } from '../services/LicenseService';
import {
  ILicenseData,
  IModuleLicenseCheck,
  PremiumModule,
  LicenseStatus,
  LicenseTier
} from '../models/ILicense';

export interface IUseLicenseResult {
  /** Current license data */
  license: ILicenseData | null;
  /** Loading state */
  loading: boolean;
  /** Error message if any */
  error: string | null;
  /** Check if a specific module is licensed */
  isModuleLicensed: (moduleId: PremiumModule) => boolean;
  /** Check multiple modules at once */
  checkModules: (moduleIds: PremiumModule[]) => IModuleLicenseCheck[];
  /** Get list of all licensed modules */
  licensedModules: PremiumModule[];
  /** Refresh license data from server */
  refresh: () => Promise<void>;
  /** Is this a trial license */
  isTrial: boolean;
  /** Is license expiring soon (within 30 days) */
  isExpiringSoon: boolean;
  /** Days until expiration (undefined if no expiration) */
  daysUntilExpiration: number | undefined;
  /** Current license tier */
  tier: LicenseTier;
  /** Is the license valid and active */
  isActive: boolean;
}

export function useLicense(licenseService: LicenseService): IUseLicenseResult {
  const [license, setLicense] = useState<ILicenseData | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const fetchLicense = useCallback(async (forceRefresh: boolean = false) => {
    setLoading(true);
    setError(null);

    try {
      const data = await licenseService.getLicense(forceRefresh);
      setLicense(data);
    } catch (err: any) {
      setError(err.message || 'Failed to fetch license data');
      console.error('Error in useLicense:', err);
    } finally {
      setLoading(false);
    }
  }, [licenseService]);

  const refresh = useCallback(async () => {
    await fetchLicense(true);
  }, [fetchLicense]);

  // Check if a specific module is licensed
  const isModuleLicensed = useCallback((moduleId: PremiumModule): boolean => {
    if (!license) return false;
    if (!license.isValid) return false;
    if (license.status !== LicenseStatus.Active && license.status !== LicenseStatus.Trial) return false;
    return license.enabledModules.includes(moduleId);
  }, [license]);

  // Check multiple modules at once
  const checkModules = useCallback((moduleIds: PremiumModule[]): IModuleLicenseCheck[] => {
    return moduleIds.map(moduleId => ({
      moduleId,
      isLicensed: isModuleLicensed(moduleId),
      reason: isModuleLicensed(moduleId) ? undefined : 'not_in_tier'
    }));
  }, [isModuleLicensed]);

  // Memoized values
  const licensedModules = useMemo(() => {
    return license?.enabledModules || [];
  }, [license]);

  const isTrial = useMemo(() => {
    return license?.isTrial || false;
  }, [license]);

  const isExpiringSoon = useMemo(() => {
    return license?.isExpiringSoon || false;
  }, [license]);

  const daysUntilExpiration = useMemo(() => {
    return license?.daysUntilExpiration;
  }, [license]);

  const tier = useMemo(() => {
    return license?.tier || LicenseTier.Free;
  }, [license]);

  const isActive = useMemo(() => {
    return license?.isValid &&
      (license.status === LicenseStatus.Active || license.status === LicenseStatus.Trial);
  }, [license]);

  // Fetch on mount
  useEffect(() => {
    fetchLicense();
  }, [fetchLicense]);

  return {
    license,
    loading,
    error,
    isModuleLicensed,
    checkModules,
    licensedModules,
    refresh,
    isTrial,
    isExpiringSoon,
    daysUntilExpiration,
    tier,
    isActive: isActive || false
  };
}

/**
 * Helper hook to check a single module
 * Use when you only need to check one module in a component
 */
export function useModuleAccess(
  licenseService: LicenseService,
  moduleId: PremiumModule
): {
  isLicensed: boolean;
  loading: boolean;
  reason?: 'not_in_tier' | 'license_expired' | 'license_suspended' | 'no_license' | 'trial_ended';
} {
  const { isModuleLicensed, loading, license } = useLicense(licenseService);

  const isLicensed = isModuleLicensed(moduleId);

  let reason: 'not_in_tier' | 'license_expired' | 'license_suspended' | 'no_license' | 'trial_ended' | undefined;

  if (!isLicensed && license) {
    if (license.status === LicenseStatus.Expired) {
      reason = 'license_expired';
    } else if (license.status === LicenseStatus.Suspended) {
      reason = 'license_suspended';
    } else if (license.tier === LicenseTier.Free && license.enabledModules.length === 0) {
      reason = 'no_license';
    } else {
      reason = 'not_in_tier';
    }
  }

  return { isLicensed, loading, reason };
}
