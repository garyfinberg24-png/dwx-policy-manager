/**
 * PnP Configuration - Standalone Policy Manager
 */
import { SPFI, spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/attachments';
import '@pnp/sp/batching';
import '@pnp/sp/search';
import '@pnp/sp/sputilities';
import '@pnp/sp/taxonomy';

let _sp: SPFI | null = null;

/**
 * Initializes PnP SP instance with the webpart context
 */
export function initializePnP(context: WebPartContext): SPFI {
  _sp = spfi().using(SPFx(context));
  return _sp;
}

/**
 * Gets the PnP SP instance. If context is provided, initializes SP first.
 */
export function getSP(context?: WebPartContext): SPFI {
  if (context) {
    return initializePnP(context);
  }
  if (!_sp) {
    throw new Error('PnP SP not initialized. Call initializePnP or getSP with context first.');
  }
  return _sp;
}

export default { initializePnP, getSP };
