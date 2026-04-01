/**
 * HealthCheckService — One-click diagnostic test suite for the Event Viewer.
 * Runs comprehensive checks against SP lists, Azure Functions, config keys,
 * and the notification queue, returning pass/fail results with remediation hints.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  PolicyLists, QuizLists, PolicyPackLists, ApprovalLists,
  NotificationLists, AdminLists,
} from '../../constants/SharePointListNames';

// ============================================================================
// TYPES
// ============================================================================

export interface IHealthCheckResult {
  name: string;
  category: 'sp-lists' | 'azure-functions' | 'configuration' | 'queue-health';
  passed: boolean;
  detail: string;
  remediation?: string;
}

export interface IHealthCheckSummary {
  results: IHealthCheckResult[];
  totalChecks: number;
  passed: number;
  failed: number;
  timestamp: string;
  durationMs: number;
}

// ============================================================================
// LISTS TO CHECK (all active PM_ lists)
// ============================================================================

const CORE_LISTS: string[] = [
  PolicyLists.POLICIES,
  PolicyLists.POLICY_VERSIONS,
  PolicyLists.POLICY_ACKNOWLEDGEMENTS,
  PolicyLists.POLICY_DISTRIBUTIONS,
  PolicyLists.POLICY_TEMPLATES,
  PolicyLists.POLICY_AUDIT_LOG,
  PolicyLists.POLICY_METADATA_PROFILES,
  PolicyLists.POLICY_REVIEWERS,
  PolicyLists.POLICY_REQUESTS,
  PolicyLists.POLICY_SUB_CATEGORIES,
  PolicyLists.POLICY_EXEMPTIONS,
  PolicyLists.POLICY_ANALYTICS,
  QuizLists.POLICY_QUIZZES,
  QuizLists.POLICY_QUIZ_QUESTIONS,
  QuizLists.POLICY_QUIZ_RESULTS,
  PolicyPackLists.POLICY_PACKS,
  PolicyPackLists.POLICY_PACK_ASSIGNMENTS,
  ApprovalLists.APPROVALS,
  ApprovalLists.APPROVAL_HISTORY,
  ApprovalLists.APPROVAL_DELEGATIONS,
  NotificationLists.NOTIFICATIONS,
  NotificationLists.NOTIFICATION_QUEUE,
  NotificationLists.REMINDER_SCHEDULE,
  AdminLists.CONFIGURATION,
  AdminLists.USER_PROFILES,
  AdminLists.EVENT_LOG,
];

// Required config keys — system won't function properly without these
const REQUIRED_CONFIG_KEYS: string[] = [
  'Admin.EventViewer.Enabled',
  'Admin.General.DefaultViewMode',
  'Admin.Compliance.RequireAcknowledgement',
  'Admin.Compliance.DefaultDeadlineDays',
  'Admin.Approval.RequireForNew',
];

// ============================================================================
// SERVICE
// ============================================================================

export class HealthCheckService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Run all health checks and return a summary.
   */
  public async runAll(): Promise<IHealthCheckSummary> {
    const startTime = Date.now();
    const results: IHealthCheckResult[] = [];

    // Run check categories in parallel
    const [listResults, configResults, queueResults] = await Promise.all([
      this._checkSpLists(),
      this._checkConfigKeys(),
      this._checkQueueHealth(),
    ]);

    results.push(...listResults, ...configResults, ...queueResults);

    // AI function checks (quick, non-blocking)
    const aiResults = await this._checkAzureFunctions();
    results.push(...aiResults);

    const passed = results.filter(r => r.passed).length;
    return {
      results,
      totalChecks: results.length,
      passed,
      failed: results.length - passed,
      timestamp: new Date().toISOString(),
      durationMs: Date.now() - startTime,
    };
  }

  // ==========================================================================
  // SP LIST CHECKS
  // ==========================================================================

  private async _checkSpLists(): Promise<IHealthCheckResult[]> {
    const results: IHealthCheckResult[] = [];

    for (const listName of CORE_LISTS) {
      try {
        const list = await this._sp.web.lists.getByTitle(listName).select('ItemCount')();
        results.push({
          name: listName,
          category: 'sp-lists',
          passed: true,
          detail: `Reachable — ${list.ItemCount} item${list.ItemCount !== 1 ? 's' : ''}`,
        });
      } catch (err: any) {
        const is404 = err?.status === 404 || err?.message?.includes('does not exist');
        results.push({
          name: listName,
          category: 'sp-lists',
          passed: false,
          detail: is404 ? 'List not found' : `Error: ${err?.message || 'Unknown'}`,
          remediation: is404
            ? `Run the provisioning script to create ${listName}. See scripts/policy-management/Deploy-AllPolicyLists.ps1`
            : 'Check SharePoint permissions or site connectivity.',
        });
      }
    }

    return results;
  }

  // ==========================================================================
  // AZURE FUNCTION CHECKS
  // ==========================================================================

  private async _checkAzureFunctions(): Promise<IHealthCheckResult[]> {
    const results: IHealthCheckResult[] = [];

    // Check AI Chat/Triage function URL from localStorage
    const chatUrl = typeof localStorage !== 'undefined'
      ? localStorage.getItem('PM_AI_ChatFunctionUrl') || ''
      : '';
    const triageUrl = typeof localStorage !== 'undefined'
      ? localStorage.getItem('PM_AI_EventTriageFunctionUrl') || ''
      : '';

    if (chatUrl) {
      results.push(await this._pingFunction('AI Chat Function', chatUrl));
    } else {
      results.push({
        name: 'AI Chat Function',
        category: 'azure-functions',
        passed: false,
        detail: 'Function URL not configured',
        remediation: 'Set the AI Chat Function URL in Admin Centre > AI Assistant settings.',
      });
    }

    if (triageUrl) {
      results.push(await this._pingFunction('AI Event Triage Function', triageUrl));
    } else {
      results.push({
        name: 'AI Event Triage Function',
        category: 'azure-functions',
        passed: false,
        detail: 'Function URL not configured',
        remediation: 'Set the Event Triage Function URL in Admin Centre > Event Viewer settings.',
      });
    }

    return results;
  }

  private async _pingFunction(name: string, url: string): Promise<IHealthCheckResult> {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000);
      // POST with empty body — Azure Functions return 400 (bad request) if reachable
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}',
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      // Any response means the function is reachable (even 400/401 is fine)
      return {
        name,
        category: 'azure-functions',
        passed: true,
        detail: `Reachable — HTTP ${response.status}`,
      };
    } catch (err: any) {
      const isTimeout = err?.name === 'AbortError';
      return {
        name,
        category: 'azure-functions',
        passed: false,
        detail: isTimeout ? 'Request timed out (10s)' : `Unreachable: ${err?.message || 'Unknown'}`,
        remediation: isTimeout
          ? 'Azure Function may be in cold start. Try again in 30 seconds.'
          : 'Check the Azure Function deployment and CORS configuration.',
      };
    }
  }

  // ==========================================================================
  // CONFIGURATION CHECKS
  // ==========================================================================

  private async _checkConfigKeys(): Promise<IHealthCheckResult[]> {
    const results: IHealthCheckResult[] = [];

    try {
      const items = await this._sp.web.lists.getByTitle('PM_Configuration')
        .items.select('ConfigKey', 'ConfigValue', 'IsActive')
        .top(200)();

      const configMap = new Map<string, { value: string; active: boolean }>();
      for (const item of items) {
        configMap.set(item.ConfigKey, { value: item.ConfigValue || '', active: item.IsActive !== false });
      }

      for (const key of REQUIRED_CONFIG_KEYS) {
        const entry = configMap.get(key);
        if (!entry) {
          results.push({
            name: key,
            category: 'configuration',
            passed: false,
            detail: 'Key not found in PM_Configuration',
            remediation: `Add ${key} to PM_Configuration list. Check Admin Centre to set this value.`,
          });
        } else if (!entry.value) {
          results.push({
            name: key,
            category: 'configuration',
            passed: false,
            detail: 'Key exists but has no value',
            remediation: `Set a value for ${key} in Admin Centre.`,
          });
        } else {
          results.push({
            name: key,
            category: 'configuration',
            passed: true,
            detail: `Value: "${entry.value.length > 40 ? entry.value.substring(0, 40) + '...' : entry.value}"`,
          });
        }
      }

      // Overall config count
      results.push({
        name: 'Total Config Keys',
        category: 'configuration',
        passed: configMap.size > 0,
        detail: `${configMap.size} key${configMap.size !== 1 ? 's' : ''} found in PM_Configuration`,
      });
    } catch (err: any) {
      results.push({
        name: 'PM_Configuration',
        category: 'configuration',
        passed: false,
        detail: `Failed to read: ${err?.message || 'Unknown'}`,
        remediation: 'Ensure PM_Configuration list exists and has correct schema.',
      });
    }

    return results;
  }

  // ==========================================================================
  // QUEUE HEALTH (DLQ / STUCK ITEMS)
  // ==========================================================================

  private async _checkQueueHealth(): Promise<IHealthCheckResult[]> {
    const results: IHealthCheckResult[] = [];
    const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000).toISOString();

    // Check notification queue for stuck/failed items
    try {
      const failedItems = await this._sp.web.lists.getByTitle('PM_NotificationQueue')
        .items
        .filter(`QueueStatus eq 'Failed'`)
        .select('Id', 'Title', 'QueueStatus')
        .top(50)();

      if (failedItems.length > 0) {
        results.push({
          name: 'Failed Notifications',
          category: 'queue-health',
          passed: false,
          detail: `${failedItems.length} failed item${failedItems.length !== 1 ? 's' : ''} in PM_NotificationQueue`,
          remediation: 'Check the Logic App run history in Azure Portal. Failed items may need manual retry or status reset.',
        });
      } else {
        results.push({
          name: 'Failed Notifications',
          category: 'queue-health',
          passed: true,
          detail: 'No failed items in notification queue',
        });
      }

      // Check for stuck "Pending" items older than 1 hour
      const stuckItems = await this._sp.web.lists.getByTitle('PM_NotificationQueue')
        .items
        .filter(`QueueStatus eq 'Pending' and Created lt '${oneHourAgo}'`)
        .select('Id', 'Title')
        .top(50)();

      if (stuckItems.length > 0) {
        results.push({
          name: 'Stuck Queue Items',
          category: 'queue-health',
          passed: false,
          detail: `${stuckItems.length} item${stuckItems.length !== 1 ? 's' : ''} stuck in "Pending" for >1 hour`,
          remediation: 'Logic App may not be polling. Check the Logic App trigger is enabled in Azure Portal.',
        });
      } else {
        results.push({
          name: 'Stuck Queue Items',
          category: 'queue-health',
          passed: true,
          detail: 'No items stuck in pending state',
        });
      }
    } catch (err: any) {
      results.push({
        name: 'Notification Queue',
        category: 'queue-health',
        passed: false,
        detail: `Failed to check: ${err?.message || 'Unknown'}`,
        remediation: 'Ensure PM_NotificationQueue list exists with QueueStatus column.',
      });
    }

    return results;
  }
}
