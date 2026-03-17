/**
 * convertDocument — Azure Function (HTTP Trigger)
 *
 * Converts .docx files to clean HTML using mammoth.js.
 * Called by PolicyService when a policy is published.
 *
 * Flow:
 *   1. Client sends: { siteUrl, documentUrl, policyId }
 *   2. Function downloads the .docx from SharePoint
 *   3. Converts to clean HTML using mammoth.js
 *   4. Returns { html, messages } to client
 *   5. Client saves HTML to PolicyContent field on PM_Policies
 *
 * Why server-side?
 *   - mammoth.js needs Node.js (can't run in SPFx browser bundle)
 *   - One-time conversion at publish, not per-read
 *   - Clean HTML stored in SP list field = instant reader rendering
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import * as mammoth from 'mammoth';
import { ConfidentialClientApplication } from '@azure/msal-node';

const TENANT_ID = process.env.AZURE_TENANT_ID || '';
const CLIENT_ID = process.env.AZURE_CLIENT_ID || '';
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || '';

interface ConvertRequest {
  siteUrl: string;
  documentUrl: string; // Server-relative URL to the .docx file
  policyId: number;
}

async function getAccessToken(siteUrl: string): Promise<string> {
  const hostname = new URL(siteUrl).hostname;
  const cca = new ConfidentialClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      clientSecret: CLIENT_SECRET
    }
  });

  const result = await cca.acquireTokenByClientCredential({
    scopes: [`https://${hostname}/.default`]
  });

  if (!result?.accessToken) throw new Error('Failed to acquire access token');
  return result.accessToken;
}

async function downloadFile(siteUrl: string, documentUrl: string, token: string): Promise<Buffer> {
  // Use SharePoint REST API to download the file
  const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(documentUrl)}')/$value`;

  const response = await fetch(apiUrl, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/octet-stream'
    }
  });

  if (!response.ok) {
    throw new Error(`Failed to download document: ${response.status} ${response.statusText}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

function applyPolicyStyles(html: string): string {
  // Wrap in a styled container that matches the Policy Manager Forest Teal theme
  return `
<div class="pm-policy-content" style="font-family: 'Segoe UI', -apple-system, sans-serif; color: #334155; line-height: 1.8; font-size: 15px; max-width: 800px; margin: 0 auto;">
  <style>
    .pm-policy-content h1 { font-size: 24px; font-weight: 700; color: #0f172a; margin: 32px 0 16px; padding-bottom: 8px; border-bottom: 2px solid #0d9488; }
    .pm-policy-content h2 { font-size: 20px; font-weight: 700; color: #0f172a; margin: 28px 0 12px; padding-bottom: 6px; border-bottom: 1px solid #e2e8f0; }
    .pm-policy-content h3 { font-size: 17px; font-weight: 600; color: #0f172a; margin: 24px 0 8px; }
    .pm-policy-content h4 { font-size: 15px; font-weight: 600; color: #334155; margin: 20px 0 8px; }
    .pm-policy-content p { margin-bottom: 12px; }
    .pm-policy-content ul, .pm-policy-content ol { margin: 8px 0 16px 24px; }
    .pm-policy-content li { margin-bottom: 6px; }
    .pm-policy-content table { border-collapse: collapse; width: 100%; margin: 16px 0; }
    .pm-policy-content th { background: #f0fdfa; color: #0f766e; font-weight: 600; text-align: left; padding: 10px 12px; border: 1px solid #e2e8f0; }
    .pm-policy-content td { padding: 10px 12px; border: 1px solid #e2e8f0; }
    .pm-policy-content tr:nth-child(even) td { background: #fafafa; }
    .pm-policy-content img { max-width: 100%; height: auto; border-radius: 4px; margin: 12px 0; }
    .pm-policy-content a { color: #0d9488; text-decoration: none; }
    .pm-policy-content a:hover { text-decoration: underline; }
    .pm-policy-content strong { color: #0f172a; }
    .pm-policy-content blockquote { border-left: 4px solid #0d9488; padding: 12px 20px; margin: 16px 0; background: #f0fdfa; color: #0f766e; font-style: italic; }
  </style>
  ${html}
</div>`;
}

export async function convertDocument(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  context.log('Document conversion request received');

  // CORS
  const origin = request.headers.get('origin') || '';
  const corsHeaders: Record<string, string> = {
    'Access-Control-Allow-Origin': origin.includes('.sharepoint.com') ? origin : '',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
  };

  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders };
  }

  try {
    const body = await request.json() as ConvertRequest;
    const { siteUrl, documentUrl, policyId } = body;

    if (!siteUrl || !documentUrl) {
      return {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: 'siteUrl and documentUrl are required' })
      };
    }

    context.log(`Converting document: ${documentUrl} for policy ${policyId}`);

    // 1. Get access token
    const token = await getAccessToken(siteUrl);

    // 2. Download the .docx file
    const fileBuffer = await downloadFile(siteUrl, documentUrl, token);
    context.log(`Downloaded ${fileBuffer.length} bytes`);

    // 3. Convert to HTML using mammoth
    const result = await mammoth.convertToHtml(
      { buffer: fileBuffer },
      {
        styleMap: [
          // Map Word styles to clean HTML
          "p[style-name='Heading 1'] => h1:fresh",
          "p[style-name='Heading 2'] => h2:fresh",
          "p[style-name='Heading 3'] => h3:fresh",
          "p[style-name='Heading 4'] => h4:fresh",
          "p[style-name='Title'] => h1.document-title:fresh",
          "p[style-name='Subtitle'] => p.document-subtitle:fresh",
          "p[style-name='Quote'] => blockquote:fresh",
          "p[style-name='Intense Quote'] => blockquote.intense:fresh",
          "r[style-name='Strong'] => strong",
          "r[style-name='Emphasis'] => em",
        ],
        convertImage: (mammoth.images as any).inline(function(element: any) {
          // Convert embedded images to base64 data URIs
          return element.read('base64').then(function(imageBuffer: string) {
            return {
              src: 'data:' + element.contentType + ';base64,' + imageBuffer
            };
          });
        })
      }
    );

    const styledHtml = applyPolicyStyles(result.value);

    context.log(`Conversion complete: ${styledHtml.length} chars, ${result.messages.length} messages`);

    // 4. Return the HTML
    return {
      status: 200,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        html: styledHtml,
        rawHtml: result.value,
        messages: result.messages.map(m => ({ type: m.type, message: m.message })),
        characterCount: styledHtml.length,
        policyId
      })
    };

  } catch (error: any) {
    context.error('Document conversion failed:', error);
    return {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        error: 'Document conversion failed',
        details: error.message || 'Unknown error'
      })
    };
  }
}

app.http('convertDocument', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'function',
  handler: convertDocument
});
