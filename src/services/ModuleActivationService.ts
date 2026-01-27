// @ts-nocheck
/**
 * ModuleActivationService - Premium Module Provisioning
 *
 * Handles the activation and provisioning of premium modules:
 * - Checks if required SharePoint lists exist
 * - Reports activation status for each module
 * - Provides instructions for manual list provisioning
 *
 * NOTE: Actual list creation is done via PowerShell scripts for security.
 * This service only checks status and provides guidance.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';

import { PremiumModule } from '../models/ILicense';
import {
  IModuleDefinition,
  IModuleActivationStatus,
  IListDefinition,
  ModuleRegistry,
  getRequiredListsForModules
} from '../models/IModuleRegistry';
import { logger } from './LoggingService';

/**
 * Result of checking a single list's existence
 */
export interface IListCheckResult {
  listName: string;
  exists: boolean;
  title?: string;
  itemCount?: number;
  errorMessage?: string;
}

/**
 * Result of module provisioning check
 */
export interface IModuleProvisioningResult {
  moduleId: PremiumModule;
  moduleName: string;
  isFullyProvisioned: boolean;
  listsChecked: IListCheckResult[];
  missingLists: IListDefinition[];
  provisioningScripts: string[];
}

/**
 * Overall provisioning status for all modules
 */
export interface IProvisioningStatus {
  timestamp: Date;
  modules: IModuleProvisioningResult[];
  allListsExist: boolean;
  missingListCount: number;
  totalListCount: number;
}

export class ModuleActivationService {
  private sp: SPFI;
  private listExistenceCache: Map<string, boolean> = new Map();
  private cacheTimestamp: number = 0;
  private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ===================================================================
  // Public Methods
  // ===================================================================

  /**
   * Get activation status for a specific module
   */
  public async getModuleActivationStatus(
    moduleId: PremiumModule,
    isLicensed: boolean
  ): Promise<IModuleActivationStatus> {
    const module = ModuleRegistry[moduleId];
    if (!module) {
      return {
        moduleId,
        isLicensed,
        isProvisioned: false,
        missingLists: ['Unknown module'],
        canActivate: false,
        activationBlockers: ['Module not found in registry']
      };
    }

    const missingLists: string[] = [];
    const activationBlockers: string[] = [];

    // Check if licensed
    if (!isLicensed) {
      activationBlockers.push('Module is not included in your current license tier');
    }

    // Check required lists
    for (const list of module.requiredLists) {
      const exists = await this.checkListExists(list.name);
      if (!exists) {
        missingLists.push(list.name);
      }
    }

    if (missingLists.length > 0) {
      activationBlockers.push(`Missing ${missingLists.length} required SharePoint list(s)`);
    }

    // Check dependencies
    if (module.dependencies) {
      for (const dep of module.dependencies) {
        const depModule = ModuleRegistry[dep.moduleId];
        const depStatus = await this.getModuleActivationStatus(dep.moduleId, isLicensed);
        if (!depStatus.isProvisioned) {
          activationBlockers.push(`Depends on ${depModule?.name || dep.moduleId} which is not provisioned`);
        }
      }
    }

    return {
      moduleId,
      isLicensed,
      isProvisioned: missingLists.length === 0,
      missingLists,
      canActivate: isLicensed && missingLists.length === 0 && activationBlockers.length <= 1,
      activationBlockers
    };
  }

  /**
   * Get provisioning status for multiple modules
   */
  public async checkModuleProvisioning(
    moduleIds: PremiumModule[]
  ): Promise<IModuleProvisioningResult[]> {
    const results: IModuleProvisioningResult[] = [];

    for (const moduleId of moduleIds) {
      const module = ModuleRegistry[moduleId];
      if (!module) continue;

      const listResults: IListCheckResult[] = [];
      const missingLists: IListDefinition[] = [];
      const provisioningScripts: string[] = [];

      for (const list of module.requiredLists) {
        const exists = await this.checkListExists(list.name);
        listResults.push({
          listName: list.name,
          exists,
          title: list.title
        });

        if (!exists) {
          missingLists.push(list);
          if (list.provisioningScript) {
            provisioningScripts.push(list.provisioningScript);
          }
        }
      }

      results.push({
        moduleId,
        moduleName: module.name,
        isFullyProvisioned: missingLists.length === 0,
        listsChecked: listResults,
        missingLists,
        provisioningScripts
      });
    }

    return results;
  }

  /**
   * Get overall provisioning status for all premium modules
   */
  public async getFullProvisioningStatus(): Promise<IProvisioningStatus> {
    const allModuleIds = Object.keys(ModuleRegistry) as PremiumModule[];
    const modules = await this.checkModuleProvisioning(allModuleIds);

    let missingListCount = 0;
    let totalListCount = 0;

    for (const module of modules) {
      totalListCount += module.listsChecked.length;
      missingListCount += module.missingLists.length;
    }

    return {
      timestamp: new Date(),
      modules,
      allListsExist: missingListCount === 0,
      missingListCount,
      totalListCount
    };
  }

  /**
   * Check if a specific list exists in SharePoint
   */
  public async checkListExists(listName: string): Promise<boolean> {
    // Check cache
    if (this.isCacheValid() && this.listExistenceCache.has(listName)) {
      return this.listExistenceCache.get(listName)!;
    }

    try {
      const list = await this.sp.web.lists.getByTitle(listName)();
      this.listExistenceCache.set(listName, true);
      this.updateCacheTimestamp();
      return true;
    } catch (error: unknown) {
      // List doesn't exist or access denied
      const errorMessage = error instanceof Error ? error.message : String(error);
      if (errorMessage.includes('does not exist') || errorMessage.includes('404')) {
        this.listExistenceCache.set(listName, false);
        this.updateCacheTimestamp();
        return false;
      }
      // Other errors - don't cache
      logger.warn('ModuleActivationService', `Error checking list ${listName}`, error);
      return false;
    }
  }

  /**
   * Get details about a specific list
   */
  public async getListDetails(listName: string): Promise<IListCheckResult> {
    try {
      const list = await this.sp.web.lists.getByTitle(listName)
        .select('Title', 'ItemCount', 'Description')();

      return {
        listName,
        exists: true,
        title: list.Title,
        itemCount: list.ItemCount
      };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return {
        listName,
        exists: false,
        errorMessage
      };
    }
  }

  /**
   * Get all required lists for a set of modules
   */
  public getRequiredLists(moduleIds: PremiumModule[]): IListDefinition[] {
    return getRequiredListsForModules(moduleIds);
  }

  /**
   * Generate provisioning script command for missing lists
   */
  public generateProvisioningCommand(missingLists: IListDefinition[]): string {
    if (missingLists.length === 0) {
      return '# All required lists are already provisioned!';
    }

    const scripts = missingLists
      .filter(list => list.provisioningScript)
      .map(list => list.provisioningScript);

    if (scripts.length === 0) {
      return '# No provisioning scripts available for missing lists';
    }

    const uniqueScripts = Array.from(new Set(scripts));

    let command = '# Run these PowerShell scripts to provision missing lists:\n';
    command += '# Make sure you are connected to SharePoint first:\n';
    command += '# Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/JML" -Interactive\n\n';

    for (const script of uniqueScripts) {
      command += `.\\scripts\\${script}\n`;
    }

    return command;
  }

  /**
   * Get module definition from registry
   */
  public getModuleDefinition(moduleId: PremiumModule): IModuleDefinition | undefined {
    return ModuleRegistry[moduleId];
  }

  /**
   * Get all module definitions
   */
  public getAllModuleDefinitions(): IModuleDefinition[] {
    return Object.values(ModuleRegistry);
  }

  /**
   * Clear the list existence cache
   */
  public clearCache(): void {
    this.listExistenceCache.clear();
    this.cacheTimestamp = 0;
  }

  // ===================================================================
  // Private Methods
  // ===================================================================

  private isCacheValid(): boolean {
    return Date.now() - this.cacheTimestamp < this.CACHE_DURATION;
  }

  private updateCacheTimestamp(): void {
    if (!this.isCacheValid()) {
      this.cacheTimestamp = Date.now();
    }
  }
}
