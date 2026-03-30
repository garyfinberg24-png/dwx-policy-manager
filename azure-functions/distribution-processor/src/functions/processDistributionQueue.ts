/**
 * processDistributionQueue — Azure Function (Timer Trigger)
 *
 * Polls PM_DistributionQueue every 2 minutes for Queued jobs.
 * Processes each job in server-side batches:
 *   1. Creates PM_PolicyAcknowledgements records
 *   2. Queues email notifications to PM_EmailQueue
 *   3. Updates progress on the queue item
 *
 * This runs SERVER-SIDE — survives browser close.
 *
 * Authentication: Uses SharePoint App-Only principal (Azure AD app registration)
 * with Sites.FullControl.All permission.
 */

import { app, Timer, InvocationContext } from '@azure/functions';
import { ConfidentialClientApplication } from '@azure/msal-node';

// --- Configuration ---
const SITE_URL = process.env.SP_SITE_URL || 'https://mf7m.sharepoint.com/sites/PolicyManager';
const TENANT_ID = process.env.AZURE_TENANT_ID || '';
const CLIENT_ID = process.env.AZURE_CLIENT_ID || '';
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || '';
const BATCH_SIZE = 50; // Users processed per batch

// --- SP List Names ---
const QUEUE_LIST = 'PM_DistributionQueue';
const ACK_LIST = 'PM_PolicyAcknowledgements';
const EMAIL_QUEUE_LIST = 'PM_NotificationQueue';
const POLICIES_LIST = 'PM_Policies';

interface QueueItem {
  Id: number;
  PolicyId: number;
  PolicyName: string;
  PolicyVersionNumber: string;
  TargetUserIds: string;
  TotalUsers: number;
  ProcessedUsers: number;
  FailedUsers: number;
  QueueStatus: string;
  JobType: string;
  DueDate?: string;
  SendNotifications: boolean;
  QueuedBy: string;
  QueuedByEmail: string;
}

// --- Get SharePoint access token ---
async function getAccessToken(): Promise<string> {
  const cca = new ConfidentialClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      clientSecret: CLIENT_SECRET
    }
  });

  const result = await cca.acquireTokenByClientCredential({
    scopes: [`https://${new URL(SITE_URL).hostname}/.default`]
  });

  if (!result?.accessToken) throw new Error('Failed to acquire access token');
  return result.accessToken;
}

// --- SharePoint REST helpers ---
async function spGet(token: string, endpoint: string): Promise<any> {
  const url = `${SITE_URL}/_api/${endpoint}`;
  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata'
    }
  });
  if (!response.ok) throw new Error(`SP GET failed: ${response.status} ${response.statusText}`);
  const data = await response.json();
  return data.value || data;
}

async function spPost(token: string, endpoint: string, body: any): Promise<any> {
  const url = `${SITE_URL}/_api/${endpoint}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': '0' // App-only doesn't need digest
    },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`SP POST failed: ${response.status} — ${errText}`);
  }
  return response.json();
}

async function spPatch(token: string, listName: string, itemId: number, body: any): Promise<void> {
  const url = `${SITE_URL}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE',
      'X-RequestDigest': '0'
    },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`SP PATCH failed: ${response.status} — ${errText}`);
  }
}

// --- Main processing function ---
async function processDistributionQueue(timer: Timer, context: InvocationContext): Promise<void> {
  context.log('Distribution queue processor triggered');

  let token: string;
  try {
    token = await getAccessToken();
  } catch (authErr) {
    context.error('Failed to acquire SP token:', authErr);
    return;
  }

  // 1. Poll for queued jobs
  let queuedJobs: QueueItem[];
  try {
    queuedJobs = await spGet(token,
      `web/lists/getbytitle('${QUEUE_LIST}')/items?$filter=QueueStatus eq 'Queued'&$top=5&$orderby=Created`
    );
  } catch (pollErr) {
    context.error('Failed to poll queue:', pollErr);
    return;
  }

  if (!queuedJobs || queuedJobs.length === 0) {
    context.log('No queued jobs found');
    return;
  }

  context.log(`Found ${queuedJobs.length} queued job(s)`);

  for (const job of queuedJobs) {
    context.log(`Processing job ${job.Id}: ${job.PolicyName} (${job.TotalUsers} users, type: ${job.JobType})`);

    try {
      // 2. Mark as Processing
      await spPatch(token, QUEUE_LIST, job.Id, {
        QueueStatus: 'Processing',
        StartedDate: new Date().toISOString()
      });

      // 3. Parse target user IDs
      let userIds: number[];
      try {
        userIds = JSON.parse(job.TargetUserIds);
      } catch {
        throw new Error('Invalid TargetUserIds JSON');
      }

      let processedCount = job.ProcessedUsers || 0;
      let failedCount = job.FailedUsers || 0;
      const errors: string[] = [];

      // 4. Process in batches
      for (let i = processedCount; i < userIds.length; i += BATCH_SIZE) {
        const batch = userIds.slice(i, i + BATCH_SIZE);

        for (const userId of batch) {
          try {
            // Check if acknowledgement already exists
            const existing = await spGet(token,
              `web/lists/getbytitle('${ACK_LIST}')/items?$filter=PolicyId eq ${job.PolicyId} and AckUserId eq ${userId}&$top=1&$select=Id`
            );

            if (!existing || existing.length === 0) {
              // Create acknowledgement record
              await spPost(token, `web/lists/getbytitle('${ACK_LIST}')/items`, {
                Title: `${job.PolicyName} - User ${userId}`,
                PolicyId: job.PolicyId,
                PolicyVersionNumber: job.PolicyVersionNumber || '1.0',
                AckUserId: userId,
                AckStatus: 'Sent',
                AssignedDate: new Date().toISOString(),
                DueDate: job.DueDate || null,
                QuizRequired: false,
                DocumentOpenCount: 0,
                TotalReadTimeSeconds: 0,
                IsDelegated: false,
                RemindersSent: 0,
                IsExempted: false,
                IsCompliant: false
              });
            }

            // Queue email notification (if enabled)
            if (job.SendNotifications) {
              try {
                // Resolve user email from SP user ID
                let userEmail = '';
                try {
                  const userInfo = await spGet(token, `web/siteusers/getbyid(${userId})?$select=Email,Title`);
                  userEmail = userInfo?.Email || '';
                } catch { /* user resolution best-effort */ }

                if (userEmail) {
                await spPost(token, `web/lists/getbytitle('${EMAIL_QUEUE_LIST}')/items`, {
                  Title: `New Policy Requires Your Acknowledgement: ${job.PolicyName}`,
                  RecipientEmail: userEmail,
                  Message: buildNotificationEmail(job),
                  NotificationType: 'policy-published',
                  Channel: 'Email',
                  PolicyId: job.PolicyId,
                  PolicyTitle: job.PolicyName,
                  Priority: 'Normal',
                  QueueStatus: 'Pending'
                });
                }
              } catch (emailErr) {
                // Non-blocking — ack record is the important part
                context.warn(`Email queue failed for user ${userId}: ${emailErr}`);
              }
            }

            processedCount++;
          } catch (userErr: any) {
            failedCount++;
            const errMsg = `User ${userId}: ${userErr.message || 'Unknown error'}`;
            errors.push(errMsg);
            context.warn(errMsg);
          }
        }

        // Update progress after each batch
        await spPatch(token, QUEUE_LIST, job.Id, {
          ProcessedUsers: processedCount,
          FailedUsers: failedCount
        });

        context.log(`  Job ${job.Id}: ${processedCount}/${userIds.length} processed (${failedCount} failed)`);
      }

      // 5. Mark as completed
      await spPatch(token, QUEUE_LIST, job.Id, {
        QueueStatus: failedCount > 0 && processedCount === 0 ? 'Failed' : 'Completed',
        ProcessedUsers: processedCount,
        FailedUsers: failedCount,
        CompletedDate: new Date().toISOString(),
        ErrorLog: JSON.stringify(errors.slice(0, 100)) // Keep last 100 errors
      });

      context.log(`Job ${job.Id} completed: ${processedCount} processed, ${failedCount} failed`);

    } catch (jobErr: any) {
      context.error(`Job ${job.Id} failed:`, jobErr);
      try {
        await spPatch(token, QUEUE_LIST, job.Id, {
          QueueStatus: 'Failed',
          ErrorLog: JSON.stringify([jobErr.message || 'Job processing failed']),
          CompletedDate: new Date().toISOString()
        });
      } catch { /* can't update — leave as Processing for manual review */ }
    }
  }

  context.log('Distribution queue processing complete');
}

// --- Email template ---
function buildNotificationEmail(job: QueueItem): string {
  const detailsUrl = `${SITE_URL}/SitePages/PolicyDetails.aspx?policyId=${job.PolicyId}`;
  const dueDateStr = job.DueDate ? new Date(job.DueDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : 'No deadline set';

  return `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;font-family:Segoe UI,sans-serif;">
      <tr><td style="background:linear-gradient(135deg,#0d9488,#0f766e);padding:24px;text-align:center;border-radius:4px 4px 0 0;">
        <h1 style="color:#fff;font-size:20px;margin:0;">New Policy Published</h1>
        <p style="color:rgba(255,255,255,0.85);font-size:14px;margin:8px 0 0;">First Digital — DWx Policy Manager</p>
      </td></tr>
      <tr><td style="padding:24px;background:#fff;border:1px solid #e2e8f0;border-top:none;">
        <p style="font-size:14px;color:#334155;">A new policy has been published and requires your acknowledgement:</p>
        <table role="presentation" width="100%" style="margin:16px 0;border:1px solid #e2e8f0;border-radius:4px;">
          <tr><td style="padding:12px;background:#f0fdfa;font-weight:600;color:#0f766e;">Policy</td><td style="padding:12px;">${job.PolicyName}</td></tr>
          <tr><td style="padding:12px;background:#f0fdfa;font-weight:600;color:#0f766e;">Due Date</td><td style="padding:12px;">${dueDateStr}</td></tr>
        </table>
        <div style="text-align:center;margin:24px 0;">
          <a href="${detailsUrl}" style="display:inline-block;padding:12px 32px;background:#0d9488;color:#fff;text-decoration:none;border-radius:4px;font-weight:600;">Read & Acknowledge</a>
        </div>
      </td></tr>
      <tr><td style="padding:16px;background:#f8fafc;text-align:center;font-size:12px;color:#94a3b8;border-radius:0 0 4px 4px;">
        First Digital — DWx Policy Manager
      </td></tr>
    </table>
  `;
}

// --- Register timer trigger (every 2 minutes) ---
app.timer('processDistributionQueue', {
  schedule: '0 */2 * * * *', // Every 2 minutes
  handler: processDistributionQueue
});
