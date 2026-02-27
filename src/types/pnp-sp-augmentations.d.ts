/**
 * PnP SP Module Augmentations
 *
 * PnP/SP v3.x is an ESM package ("type": "module") that uses .js extensions
 * in its type declarations for module augmentations. TypeScript 4.7 with
 * moduleResolution: "node" may not always resolve these augmentations correctly.
 *
 * This file explicitly re-declares the augmentations that PnP/SP provides
 * via side-effect imports (e.g., import '@pnp/sp/lists'), ensuring that
 * IWeb has the expected properties like .lists, .siteUsers, etc.
 *
 * Reference: @pnp/sp/lists/web.d.ts, @pnp/sp/items/list.d.ts,
 *            @pnp/sp/site-users/web.d.ts
 */

import { ILists, IList } from '@pnp/sp/lists';
import { IItems } from '@pnp/sp/items';
import { ISiteUsers, ISiteUser } from '@pnp/sp/site-users';
import { ISPCollection } from '@pnp/sp/spqueryable';

// Augment IWeb with lists property (from @pnp/sp/lists/web.d.ts)
declare module '@pnp/sp/webs' {
  interface IWeb {
    readonly lists: ILists;
    readonly siteUserInfoList: IList;
    readonly defaultDocumentLibrary: IList;
    readonly customListTemplates: ISPCollection;
    getList(listRelativeUrl: string): IList;
    getCatalog(type: number): Promise<IList>;
    readonly siteUsers: ISiteUsers;
    readonly currentUser: ISiteUser;
    getUserById(id: number): ISiteUser;
    ensureUser(loginName: string): Promise<{ data: { Id: number; Title: string; LoginName: string; Email: string } }>;
  }
}

// Augment IList with items property (from @pnp/sp/items/list.d.ts)
declare module '@pnp/sp/lists' {
  interface IList {
    readonly items: IItems;
  }
}
