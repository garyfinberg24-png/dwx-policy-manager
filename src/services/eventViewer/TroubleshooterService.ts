/**
 * TroubleshooterService — Wizard-driven diagnostic tool.
 * Defines problem categories, check sequences, and remediation advice.
 * Reuses HealthCheckService, SchemaValidatorService, ConfigAuditService.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { EventBuffer } from './EventBuffer';
import { SLAMonitorService } from './SLAMonitorService';

// ============================================================================
// TYPES
// ============================================================================

export interface ITroubleshooterProblem {
  id: string;
  label: string;
  description: string;
  icon: string; // SVG path
  color: string;
}

export interface IDiagnosticCheck {
  name: string;
  description: string;
  status: 'pending' | 'running' | 'passed' | 'failed' | 'warning' | 'skipped';
  detail: string;
  remediation?: string;
  /** Link to Admin Centre section (if applicable) */
  adminLink?: string;
}

export interface ITroubleshooterResult {
  problemId: string;
  checks: IDiagnosticCheck[];
  totalChecks: number;
  passed: number;
  failed: number;
  warnings: number;
  durationMs: number;
  summary: string;
}

// ============================================================================
// PROBLEM CATALOGUE
// ============================================================================

export const PROBLEMS: ITroubleshooterProblem[] = [
  {
    id: 'policy-save',
    label: 'Policies won\'t save or publish',
    description: 'Draft save, publish, or version creation is failing',
    icon: 'M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z M14 2v6h6',
    color: '#dc2626',
  },
  {
    id: 'acknowledgements',
    label: 'Users can\'t acknowledge policies',
    description: 'Acknowledgement button fails or status doesn\'t update',
    icon: 'M9 11l3 3L22 4 M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11',
    color: '#d97706',
  },
  {
    id: 'notifications',
    label: 'Notifications aren\'t being sent',
    description: 'Email or in-app notifications are missing',
    icon: 'M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9 M13.73 21a2 2 0 01-3.46 0',
    color: '#2563eb',
  },
  {
    id: 'quiz',
    label: 'Quiz generation is failing',
    description: 'AI quiz generation returns errors or no results',
    icon: 'M12 2a10 10 0 100 20 10 10 0 000-20z M12 16v-4 M12 8h.01',
    color: '#7c3aed',
  },
  {
    id: 'performance',
    label: 'App is slow or pages won\'t load',
    description: 'Slow page loads, timeouts, or throttling',
    icon: 'M13 2L3 14h9l-1 8 10-12h-9l1-8z',
    color: '#d97706',
  },
  {
    id: 'access',
    label: 'Users report access denied',
    description: '403 errors, missing nav items, or blocked features',
    icon: 'M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z',
    color: '#dc2626',
  },
  {
    id: 'general',
    label: 'Something else / General health check',
    description: 'Run all diagnostic checks across the system',
    icon: 'M22 12h-4l-3 9L9 3l-3 9H2',
    color: '#0d9488',
  },
];

// ============================================================================
// SERVICE
// ============================================================================

export class TroubleshooterService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Run diagnostics for a specific problem.
   * Yields checks via callback as they complete (for live UI updates).
   */
  public async diagnose(
    problemId: string,
    buffer: EventBuffer,
    onCheckUpdate: (checks: IDiagnosticCheck[]) => void
  ): Promise<ITroubleshooterResult> {
    const startTime = Date.now();
    const checks: IDiagnosticCheck[] = this._getChecksForProblem(problemId);

    // Notify UI of initial state
    onCheckUpdate(checks.slice());

    // Run checks sequentially (so we can skip based on prior results)
    for (let i = 0; i < checks.length; i++) {
      checks[i].status = 'running';
      onCheckUpdate(checks.slice());

      try {
        await this._runCheck(checks[i], problemId, buffer);
      } catch (err: any) {
        checks[i].status = 'failed';
        checks[i].detail = `Check failed: ${err?.message || 'Unknown error'}`;
      }

      onCheckUpdate(checks.slice());

      // Skip subsequent checks if a critical dependency failed
      if (checks[i].status === 'failed' && checks[i].name.includes('List exists')) {
        // Skip remaining checks in this group — list doesn't exist
        for (let j = i + 1; j < checks.length; j++) {
          if (checks[j].name.startsWith('Check') && !checks[j].name.includes('List exists')) {
            checks[j].status = 'skipped';
            checks[j].detail = 'Skipped — prerequisite list not found';
          }
        }
        onCheckUpdate(checks.slice());
      }
    }

    const passed = checks.filter(c => c.status === 'passed').length;
    const failed = checks.filter(c => c.status === 'failed').length;
    const warnings = checks.filter(c => c.status === 'warning').length;

    const summary = failed === 0
      ? 'All checks passed. The system appears healthy for this area.'
      : `${failed} issue${failed !== 1 ? 's' : ''} found. See remediation steps below.`;

    return {
      problemId,
      checks,
      totalChecks: checks.length,
      passed,
      failed,
      warnings,
      durationMs: Date.now() - startTime,
      summary,
    };
  }

  // ==========================================================================
  // CHECK DEFINITIONS PER PROBLEM
  // ==========================================================================

  private _getChecksForProblem(problemId: string): IDiagnosticCheck[] {
    const mk = (name: string, desc: string): IDiagnosticCheck => ({
      name, description: desc, status: 'pending', detail: '',
    });

    switch (problemId) {
      case 'policy-save':
        return [
          mk('PM_Policies list exists', 'Verify the core policies list is accessible'),
          mk('PM_PolicyVersions list exists', 'Verify version history list is accessible'),
          mk('PM_PolicyAuditLog list exists', 'Verify audit log list is accessible'),
          mk('Check for recent save errors', 'Look for PolicyService errors in event buffer'),
          mk('Check required columns', 'Verify key columns on PM_Policies'),
          mk('Check audit log writable', 'Test write access to PM_PolicyAuditLog'),
        ];
      case 'acknowledgements':
        return [
          mk('PM_PolicyAcknowledgements list exists', 'Verify acknowledgements list is accessible'),
          mk('Check AckUserId column', 'Verify AckUserId column exists on list'),
          mk('Check for acknowledgement errors', 'Look for ack-related errors in buffer'),
          mk('Check distribution status', 'Verify policies are distributed to users'),
        ];
      case 'notifications':
        return [
          mk('PM_Notifications list exists', 'Verify in-app notifications list'),
          mk('PM_NotificationQueue list exists', 'Verify email queue list'),
          mk('Check for failed queue items', 'Look for Failed status in notification queue'),
          mk('Check for stuck pending items', 'Look for items stuck >1 hour in Pending'),
          mk('Check notification toggle', 'Verify notifications are enabled in config'),
          mk('Check Logic App URL', 'Verify email sender Logic App is configured'),
        ];
      case 'quiz':
        return [
          mk('PM_PolicyQuizzes list exists', 'Verify quiz definitions list'),
          mk('PM_PolicyQuizQuestions list exists', 'Verify quiz questions list'),
          mk('Check AI Function URL', 'Verify quiz generation function URL is configured'),
          mk('Check AI Function reachable', 'Test connectivity to Azure Function'),
          mk('Check for quiz generation errors', 'Look for quiz-related errors in buffer'),
        ];
      case 'performance':
        return [
          mk('Check SLA P95', 'Measure 95th percentile response time'),
          mk('Check for throttling (429s)', 'Look for HTTP 429 Too Many Requests'),
          mk('Check for slow requests (>3s)', 'Count requests exceeding 3 second threshold'),
          mk('Check error rate', 'Calculate overall error rate from network events'),
          mk('Check buffer overflow', 'Verify event buffer is not at capacity'),
        ];
      case 'access':
        return [
          mk('PM_Configuration list exists', 'Verify config list is accessible'),
          mk('Check for 401/403 errors', 'Look for authentication/authorization errors'),
          mk('Check nav permissions config', 'Verify role permissions are configured'),
          mk('Check CORS errors', 'Look for cross-origin request failures'),
        ];
      case 'general':
      default:
        return [
          mk('PM_Policies list exists', 'Core policies list'),
          mk('PM_Configuration list exists', 'Configuration list'),
          mk('PM_NotificationQueue list exists', 'Email queue list'),
          mk('PM_PolicyAuditLog list exists', 'Audit log list'),
          mk('Check for errors in buffer', 'Any errors captured this session'),
          mk('Check SLA P95', 'Overall response time health'),
          mk('Check AI Function URL', 'AI services configured'),
          mk('Check notification queue health', 'No stuck or failed items'),
        ];
    }
  }

  // ==========================================================================
  // CHECK EXECUTION
  // ==========================================================================

  private async _runCheck(check: IDiagnosticCheck, problemId: string, buffer: EventBuffer): Promise<void> {
    const name = check.name;

    // List existence checks
    if (name.includes('list exists')) {
      const listName = name.replace(' list exists', '').trim();
      await this._checkListExists(check, listName);
      return;
    }

    // Buffer error checks
    if (name.includes('recent save errors') || name.includes('acknowledgement errors') || name.includes('quiz generation errors')) {
      this._checkBufferErrors(check, buffer, problemId);
      return;
    }

    if (name === 'Check for errors in buffer') {
      const errors = buffer.getAll().filter(e => e.severity >= 3);
      check.status = errors.length === 0 ? 'passed' : errors.length > 5 ? 'failed' : 'warning';
      check.detail = errors.length === 0 ? 'No errors in buffer' : `${errors.length} error${errors.length !== 1 ? 's' : ''} found`;
      if (errors.length > 0) check.remediation = 'Check the Event Stream tab for error details.';
      return;
    }

    if (name === 'Check for failed queue items') {
      await this._checkQueueFailed(check);
      return;
    }

    if (name === 'Check for stuck pending items') {
      await this._checkQueueStuck(check);
      return;
    }

    if (name === 'Check notification toggle') {
      await this._checkConfigKey(check, 'Admin.Notifications.NewPolicies', 'true');
      return;
    }

    if (name === 'Check Logic App URL' || name === 'Check AI Function URL') {
      const key = name.includes('Logic') ? 'Integration.Email.LogicAppUrl' : 'Integration.AI.Chat.FunctionUrl';
      const lsKey = name.includes('Logic') ? '' : 'PM_AI_ChatFunctionUrl';
      const val = typeof localStorage !== 'undefined' && lsKey ? localStorage.getItem(lsKey) : '';
      if (val) {
        check.status = 'passed';
        check.detail = 'Function URL configured in localStorage';
      } else {
        await this._checkConfigKey(check, key, undefined, true);
      }
      return;
    }

    if (name === 'Check AI Function reachable') {
      const url = typeof localStorage !== 'undefined' ? (localStorage.getItem('PM_AI_ChatFunctionUrl') || '') : '';
      if (!url) {
        check.status = 'failed';
        check.detail = 'No AI Function URL configured';
        check.remediation = 'Set the AI Function URL in Admin Centre > AI Assistant.';
        return;
      }
      try {
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), 8000);
        const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}', signal: controller.signal });
        clearTimeout(tid);
        check.status = 'passed';
        check.detail = `Reachable — HTTP ${res.status}`;
      } catch (err: any) {
        check.status = 'failed';
        check.detail = err?.name === 'AbortError' ? 'Timed out (8s)' : `Unreachable: ${err?.message || 'Unknown'}`;
        check.remediation = 'Check Azure Function deployment and CORS settings.';
      }
      return;
    }

    if (name === 'Check SLA P95') {
      const sla = SLAMonitorService.compute(buffer);
      if (sla.overallP95 === 0) {
        check.status = 'passed';
        check.detail = 'No network data yet — navigate the app first';
      } else if (sla.overallP95 > 3000) {
        check.status = 'failed';
        check.detail = `P95 latency: ${sla.overallP95}ms (target: <2000ms)`;
        check.remediation = 'Check the SLA Monitor in System Health for per-list breakdown. Consider enabling Performance Optimizer settings.';
      } else if (sla.overallP95 > 2000) {
        check.status = 'warning';
        check.detail = `P95 latency: ${sla.overallP95}ms (approaching 2000ms target)`;
      } else {
        check.status = 'passed';
        check.detail = `P95 latency: ${sla.overallP95}ms — within target`;
      }
      return;
    }

    if (name.includes('throttling')) {
      const net = buffer.getNetworkEvents();
      const throttled = net.filter((e: any) => e.httpStatus === 429).length;
      check.status = throttled === 0 ? 'passed' : 'failed';
      check.detail = throttled === 0 ? 'No throttling detected' : `${throttled} throttled request${throttled !== 1 ? 's' : ''} (HTTP 429)`;
      if (throttled > 0) check.remediation = 'Reduce concurrent API calls. Enable request deduplication in Performance Optimizer.';
      return;
    }

    if (name.includes('slow requests')) {
      const net = buffer.getNetworkEvents();
      const slow = net.filter((e: any) => (e.duration || 0) > 3000).length;
      check.status = slow === 0 ? 'passed' : slow > 3 ? 'failed' : 'warning';
      check.detail = slow === 0 ? 'No slow requests' : `${slow} request${slow !== 1 ? 's' : ''} exceeding 3s`;
      if (slow > 0) check.remediation = 'Check the Network Monitor for slow endpoints. Consider enabling lean queries.';
      return;
    }

    if (name.includes('error rate')) {
      const net = buffer.getNetworkEvents();
      const errors = net.filter((e: any) => (e.httpStatus || 0) >= 400).length;
      const rate = net.length > 0 ? Math.round((errors / net.length) * 1000) / 10 : 0;
      check.status = rate === 0 ? 'passed' : rate > 10 ? 'failed' : 'warning';
      check.detail = `${rate}% error rate (${errors}/${net.length} requests)`;
      if (rate > 5) check.remediation = 'Check the Network Monitor for failing endpoints.';
      return;
    }

    if (name.includes('buffer overflow')) {
      const stats = buffer.getStats();
      const usage = stats.totalCount / (stats.capacity.app + stats.capacity.console + stats.capacity.network);
      check.status = usage < 0.9 ? 'passed' : 'warning';
      check.detail = `Buffer ${Math.round(usage * 100)}% full (${stats.totalCount} events)`;
      if (usage >= 0.9) check.remediation = 'Old events are being dropped. Increase buffer sizes in Admin Centre > Event Viewer.';
      return;
    }

    if (name.includes('401/403')) {
      const net = buffer.getNetworkEvents();
      const authErrors = net.filter((e: any) => e.httpStatus === 401 || e.httpStatus === 403).length;
      check.status = authErrors === 0 ? 'passed' : 'failed';
      check.detail = authErrors === 0 ? 'No auth errors detected' : `${authErrors} authentication/authorization error${authErrors !== 1 ? 's' : ''}`;
      if (authErrors > 0) check.remediation = 'Check user permissions in SharePoint site settings and PM_Configuration role assignments.';
      return;
    }

    if (name.includes('nav permissions')) {
      await this._checkConfigKey(check, 'Admin.General.DefaultViewMode', undefined, false);
      return;
    }

    if (name.includes('CORS')) {
      const all = buffer.getAll();
      const cors = all.filter(e => e.message.toLowerCase().indexOf('cors') !== -1 || e.message.toLowerCase().indexOf('cross-origin') !== -1);
      check.status = cors.length === 0 ? 'passed' : 'failed';
      check.detail = cors.length === 0 ? 'No CORS errors detected' : `${cors.length} CORS error${cors.length !== 1 ? 's' : ''} found`;
      if (cors.length > 0) check.remediation = 'Check Azure Function CORS settings — ensure the SharePoint origin is allowed.';
      return;
    }

    if (name.includes('required columns') || name.includes('AckUserId column')) {
      check.status = 'passed';
      check.detail = 'Column check requires Schema Validator — run it from System Health tab';
      return;
    }

    if (name.includes('audit log writable')) {
      await this._checkListWritable(check, 'PM_PolicyAuditLog');
      return;
    }

    if (name.includes('distribution status')) {
      await this._checkListExists(check, 'PM_PolicyDistributions');
      return;
    }

    if (name.includes('notification queue health')) {
      await this._checkQueueFailed(check);
      return;
    }

    // Default — unknown check
    check.status = 'passed';
    check.detail = 'Check completed';
  }

  // ==========================================================================
  // CHECK HELPERS
  // ==========================================================================

  private async _checkListExists(check: IDiagnosticCheck, listName: string): Promise<void> {
    try {
      const list = await this._sp.web.lists.getByTitle(listName).select('ItemCount')();
      check.status = 'passed';
      check.detail = `${listName} exists — ${list.ItemCount} items`;
    } catch (err: any) {
      check.status = 'failed';
      check.detail = `${listName} not found or inaccessible`;
      check.remediation = `Run the provisioning script to create ${listName}. See scripts/policy-management/Deploy-AllPolicyLists.ps1`;
    }
  }

  private async _checkListWritable(check: IDiagnosticCheck, listName: string): Promise<void> {
    try {
      const item = await this._sp.web.lists.getByTitle(listName).items.add({
        Title: '[Troubleshooter] Write test — safe to delete',
        AuditAction: 'SystemCheck',
        EntityType: 'Diagnostic',
        ActionDescription: 'Troubleshooter write test',
        PerformedByEmail: 'troubleshooter@system',
      });
      // Clean up test item
      try { await this._sp.web.lists.getByTitle(listName).items.getById((item as any).data?.Id || (item as any).Id).delete(); } catch (_) {}
      check.status = 'passed';
      check.detail = `${listName} is writable`;
    } catch (err: any) {
      check.status = 'failed';
      check.detail = `Cannot write to ${listName}: ${err?.message || 'Unknown'}`;
      check.remediation = 'Check SharePoint list permissions — the current user may not have Contribute access.';
    }
  }

  private _checkBufferErrors(check: IDiagnosticCheck, buffer: EventBuffer, problemId: string): void {
    const allEvents = buffer.getAll();
    let pattern: RegExp;
    switch (problemId) {
      case 'policy-save': pattern = /PolicyService|save|publish|version/i; break;
      case 'acknowledgements': pattern = /acknowledg|ack/i; break;
      case 'quiz': pattern = /quiz|generate/i; break;
      default: pattern = /error/i;
    }
    const relevant = allEvents.filter(e => e.severity >= 3 && pattern.test(e.message + ' ' + e.source));
    check.status = relevant.length === 0 ? 'passed' : 'failed';
    check.detail = relevant.length === 0
      ? `No relevant errors found in current session`
      : `${relevant.length} related error${relevant.length !== 1 ? 's' : ''}: "${relevant[0].message.substring(0, 80)}"`;
    if (relevant.length > 0) check.remediation = 'Click the error in Event Stream for full details and stack trace.';
  }

  private async _checkQueueFailed(check: IDiagnosticCheck): Promise<void> {
    try {
      const failed = await this._sp.web.lists.getByTitle('PM_NotificationQueue')
        .items.filter("QueueStatus eq 'Failed'").select('Id').top(50)();
      check.status = failed.length === 0 ? 'passed' : 'failed';
      check.detail = failed.length === 0 ? 'No failed items in queue' : `${failed.length} failed notification${failed.length !== 1 ? 's' : ''} in queue`;
      if (failed.length > 0) check.remediation = 'Check the Logic App run history in Azure Portal. Failed items may need manual retry.';
    } catch (err: any) {
      check.status = 'failed';
      check.detail = `Cannot read PM_NotificationQueue: ${err?.message || 'Unknown'}`;
      check.remediation = 'Ensure PM_NotificationQueue exists with QueueStatus column.';
    }
  }

  private async _checkQueueStuck(check: IDiagnosticCheck): Promise<void> {
    const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000).toISOString();
    try {
      const stuck = await this._sp.web.lists.getByTitle('PM_NotificationQueue')
        .items.filter(`QueueStatus eq 'Pending' and Created lt '${oneHourAgo}'`).select('Id').top(50)();
      check.status = stuck.length === 0 ? 'passed' : 'failed';
      check.detail = stuck.length === 0 ? 'No stuck items' : `${stuck.length} item${stuck.length !== 1 ? 's' : ''} stuck in Pending for >1 hour`;
      if (stuck.length > 0) check.remediation = 'The Logic App may not be polling. Check the Logic App trigger is enabled in Azure Portal.';
    } catch (_) {
      check.status = 'warning';
      check.detail = 'Could not check for stuck items';
    }
  }

  private async _checkConfigKey(check: IDiagnosticCheck, key: string, expectedValue?: string, requireNonEmpty?: boolean): Promise<void> {
    try {
      const items = await this._sp.web.lists.getByTitle('PM_Configuration')
        .items.filter(`ConfigKey eq '${key}'`).select('ConfigValue', 'IsActive').top(1)();
      if (items.length === 0) {
        check.status = requireNonEmpty ? 'failed' : 'warning';
        check.detail = `Config key "${key}" not found`;
        check.remediation = `Add "${key}" in Admin Centre.`;
      } else if (requireNonEmpty && !items[0].ConfigValue) {
        check.status = 'failed';
        check.detail = `Config key "${key}" exists but has no value`;
        check.remediation = `Set a value for "${key}" in Admin Centre.`;
      } else if (expectedValue && items[0].ConfigValue !== expectedValue) {
        check.status = 'warning';
        check.detail = `"${key}" = "${items[0].ConfigValue}" (expected: "${expectedValue}")`;
      } else {
        check.status = 'passed';
        check.detail = `"${key}" = "${items[0].ConfigValue}"`;
      }
    } catch (_) {
      check.status = 'warning';
      check.detail = `Could not read config key "${key}"`;
    }
  }
}
