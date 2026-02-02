import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
// Support either @pnp/sp v2/v3 import shapes
// Prefer the v3 'spfi' + 'SPFx' initialization when available.
let spInstance: any = null;
try {
    // Try to lazy-import v3 helpers if present
    // require is used so build doesn't fail if package shape differs
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const pnp = require('@pnp/sp');
    spInstance = (pnp as any).sp ?? (pnp as any).spfi ?? pnp;
} catch (e) {
    // fallback: package may not expose expected shapes yet — will initialize in init()
    spInstance = null;
}

export interface IPropertyItem {
    Id?: number;
    Title?: string;
    [key: string]: any;
}

export class PropertyService {
    private static _context: WebPartContext | null = null;
    /** Initialize PnPJS with SPFx context */
    public static init(context: WebPartContext) {
        // store SPFx context for REST fallback
        PropertyService._context = context;
        // Initialize PnPJS with SPFx context.
        // Support both v2 (sp.setup) and v3 (spfi().using(SPFx(context))).
        try {
            // Try v3 style initialization
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const { spfi, SPFx } = require('@pnp/sp');
            spInstance = spfi().using(SPFx(context));
            return;
        } catch (e) {
            // Not v3 shape — try v2 setup
        }

        if (spInstance && typeof spInstance.setup === 'function') {
            spInstance.setup({ spfxContext: context });
            return;
        }

        // Last resort: attempt to require and setup sp directly
        try {
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const pnp = require('@pnp/sp');
            if (pnp && typeof pnp.sp !== 'undefined' && typeof pnp.sp.setup === 'function') {
                (pnp.sp as any).setup({ spfxContext: context });
                spInstance = pnp.sp;
            }
        } catch (e) {
            console.warn('PropertyService: could not initialize PnPJS automatically', e);
        }
    }

    private static getList() {
        // Ensure PnPJS has been initialized
        if (!spInstance) {
            const msg = 'PropertyService: PnPJS not initialized. Call PropertyService.init(context) before using the service.';
            console.error(msg);
            throw new Error(msg);
        }

        // Support multiple PnPJS shapes:
        // - v2: sp.web.lists
        // - v3: sp.web.lists (should also exist)
        // - some packaging shapes may expose lists at top-level: sp.lists
        try {
            if (spInstance.web && spInstance.web.lists) {
                return spInstance.web.lists.getByTitle('Property');
            }

            if (spInstance.lists && typeof spInstance.lists.getByTitle === 'function') {
                return spInstance.lists.getByTitle('Property');
            }

            // As a last attempt, if spInstance has a 'get' function or appears to be a factory, try calling it
            if (typeof spInstance === 'function') {
                try {
                    const inst = spInstance();
                    if (inst.web && inst.web.lists) {
                        return inst.web.lists.getByTitle('Property');
                    }
                } catch (e) {
                    // ignore
                }
            }

            console.error('PropertyService: spInstance present but no .web.lists or top-level .lists found. spInstance keys=', Object.keys(spInstance));
            throw new Error('PropertyService: spInstance does not expose lists; initialization shape may be incorrect.');
        } catch (err) {
            console.error('PropertyService.getList detection error', err);
            throw err;
        }
    }

    /** Get list fields (schema) - returns useful field properties */
    public static async getFields(): Promise<any[]> {
        try {
            if (spInstance && (spInstance.web || spInstance.lists)) {
                // request additional metadata useful for detection and mapping
                const fld = await spInstance.web.lists.getByTitle('Property').fields.select(
                    'Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'ReadOnlyField', 'Hidden', 'AllowMultipleValues', 'UserSelectionMode', 'SchemaXml'
                ).get();
                return fld;
            }

            // REST fallback
            return await this.fetchFieldsViaRest();
        } catch (e) {
            console.error('PropertyService.getFields error', e);
            throw e;
        }
    }

    /** Get list views and their fields */
    public static async getViews(): Promise<any[]> {
        try {
            if (spInstance && (spInstance.web || spInstance.lists)) {
                // request more properties and a large top to include all views
                const views = await spInstance.web.lists.getByTitle('Property').views.top(5000).select('Id', 'Title', 'ViewFields', 'RowLimit', 'Hidden', 'PersonalView').get();
                return views;
            }

            // REST fallback
            return await this.fetchViewsViaRest();
        } catch (e) {
            console.error('PropertyService.getViews error', e);
            throw e;
        }
    }

    public static async getItems(filter?: string, top = 200): Promise<IPropertyItem[]> {
        try {
            if (spInstance && (spInstance.web || spInstance.lists)) {
                const q = this.getList().items.select('*').top(top);
                const items = await q.get();
                return items as IPropertyItem[];
            }

            return await this.fetchItemsViaRest(top);
        } catch (e) {
            console.error('PropertyService.getItems error', e);
            throw e;
        }
    }

    /** Get choices or lookup info for a specific field (Choice / Lookup) */
    public static async getFieldChoices(internalName: string): Promise<any> {
        try {
            if (spInstance && (spInstance.web || spInstance.lists)) {
                // PnPJS: get field by internal name or title and request Choices/LookupList/LookupField/AllowMultipleValues
                const fld = await this.getList().fields.getByInternalNameOrTitle(internalName).select('Choices', 'LookupList', 'LookupField', 'AllowMultipleValues', 'TypeAsString').get();
                return fld;
            }

            // REST fallback
            if (!PropertyService._context) throw new Error('PropertyService: no SPFx context for REST calls');
            const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/lists/getByTitle('Property')/fields/getByInternalNameOrTitle('${encodeURIComponent(internalName)}')?$select=Choices,LookupList,LookupField,AllowMultipleValues,TypeAsString`;
            // eslint-disable-next-line no-console
            console.debug('PropertyService.getFieldChoices/rest url', url);
            const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            const data = await res.json();
            // Normalize REST shapes. SharePoint sometimes wraps single objects in { d: { ... } }
            // or returns { value: [ ... ] } depending on OData settings. Normalize to a plain object
            // so callers can safely access properties like LookupList.
            let normalized: any = data;
            if (data && typeof data === 'object') {
                if ((data as any).d) normalized = (data as any).d;
                if (normalized && normalized.value && Array.isArray(normalized.value) && normalized.value.length === 1) {
                    normalized = normalized.value[0];
                }
            }
            return normalized;
        } catch (e) {
            console.warn('PropertyService.getFieldChoices failed', e);
            return null;
        }
    }

    /** Fetch items from a lookup list by list id (GUID) */
    public static async getLookupItems(lookupListId: string, top = 500): Promise<any[]> {
        try {
            if (!lookupListId) return [];
            if (spInstance && (spInstance.web || spInstance.lists)) {
                // lookupListId from field metadata is often a GUID without braces
                const list = spInstance.web.lists.getById(lookupListId);
                const items = await list.items.select('Id', 'Title').top(top).get();
                return items || [];
            }

            if (!PropertyService._context) throw new Error('PropertyService: no SPFx context for REST calls');
            const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
            // REST expects GUID in parens: lists(guid'...')
            const url = `${webUrl}/_api/web/lists(guid'${lookupListId}')/items?$select=Id,Title&$top=${top}`;
            // eslint-disable-next-line no-console
            console.debug('PropertyService.getLookupItems/rest url', url);
            const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            const data = await res.json();
            return data && data.value ? data.value : [];
        } catch (e) {
            console.warn('PropertyService.getLookupItems failed', e);
            return [];
        }
    }

    public static async getItem(id: number): Promise<IPropertyItem | null> {
        try {
            const item = await this.getList().items.getById(id).get();
            return item as IPropertyItem;
        } catch (e) {
            console.error('PropertyService.getItem error', e);
            throw e;
        }
    }

    /** Get a single user by id (returns minimal user object) */
    public static async getUserById(id: number): Promise<any | null> {
        try {
            if (!id) return null;
            // Try PnPJS if available
            if (spInstance && (spInstance.web || spInstance.siteUsers)) {
                try {
                    if (spInstance.web && spInstance.web.siteUsers && typeof spInstance.web.siteUsers.getById === 'function') {
                        const u = await spInstance.web.siteUsers.getById(id).get();
                        return { id: u.Id ?? u.ID ?? u.id, title: u.Title ?? u.title, login: u.LoginName ?? u.LoginName, email: u.Email ?? u.Email };
                    }
                } catch (e) {
                    // fall through to REST fallback
                }
            }

            if (!PropertyService._context) throw new Error('PropertyService: no SPFx context for REST calls');
            const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/getUserById(${id})?$select=Id,Title,LoginName,Email`;
            // eslint-disable-next-line no-console
            console.debug('PropertyService.getUserById/rest url', url);
            const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            const data = await res.json();
            let normalized: any = data;
            if (data && (data as any).d) normalized = (data as any).d;
            if (normalized && (normalized.Id || normalized.ID)) {
                return { id: normalized.Id ?? normalized.ID, title: normalized.Title, login: normalized.LoginName, email: normalized.Email };
            }
            return null;
        } catch (e) {
            console.warn('PropertyService.getUserById failed', e);
            return null;
        }
    }

    public static async createItem(data: { [k: string]: any }): Promise<IPropertyItem> {
        try {
            const res = await this.getList().items.add(data);
            return res.data as IPropertyItem;
        } catch (e) {
            console.error('PropertyService.createItem error', e);
            throw e;
        }
    }

    public static async updateItem(id: number, data: { [k: string]: any }): Promise<void> {
        try {
            await this.getList().items.getById(id).update(data);
        } catch (e) {
            console.error('PropertyService.updateItem error', e);
            throw e;
        }
    }

    public static async deleteItem(id: number): Promise<void> {
        try {
            await this.getList().items.getById(id).delete();
        } catch (e) {
            console.error('PropertyService.deleteItem error', e);
            throw e;
        }
    }

    // ---------- REST fallback helpers ----------
    private static async fetchFieldsViaRest(): Promise<any[]> {
        if (!PropertyService._context) {
            throw new Error('PropertyService: no SPFx context for REST calls');
        }
        const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
        const candidates = ['Property', 'Properties', 'property', 'properties'];
        for (const t of candidates) {
            // request additional metadata fields that help detect people/date/multi-value
            const url = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(t)}')/fields?$select=Id,Title,InternalName,TypeAsString,FieldTypeKind,ReadOnlyField,Hidden,AllowMultipleValues,UserSelectionMode,SchemaXml`;
            try {
                // eslint-disable-next-line no-console
                console.debug(`PropertyService.fetchFieldsViaRest: trying list title='${t}' url=${url}`);
                const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
                // eslint-disable-next-line no-console
                console.debug('PropertyService.fetchFieldsViaRest: response status', res.status);
                const data = await res.json();
                // eslint-disable-next-line no-console
                console.debug('PropertyService.fetchFieldsViaRest: response body', data);
                if (data && Array.isArray(data.value) && data.value.length > 0) {
                    // eslint-disable-next-line no-console
                    console.debug(`PropertyService.fetchFieldsViaRest: found ${data.value.length} fields for list '${t}'`);
                    return data.value;
                }
            } catch (e) {
                // eslint-disable-next-line no-console
                console.warn(`PropertyService.fetchFieldsViaRest: attempt for '${t}' failed`, e);
            }
        }

        // If none succeeded, try a generic lists lookup to help debugging
        try {
            const listUrl = `${webUrl}/_api/web/lists?$select=Title,Id,BaseTemplate,Hidden`;
            const listRes = await PropertyService._context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            const listData = await listRes.json();
            // eslint-disable-next-line no-console
            console.debug('PropertyService.fetchFieldsViaRest: lists on web', listData.value && listData.value.length ? listData.value.slice(0, 20) : listData.value);
        } catch (e) {
            // eslint-disable-next-line no-console
            console.warn('PropertyService.fetchFieldsViaRest: failed to list lists for debugging', e);
        }

        return [];
    }

    private static async fetchViewsViaRest(): Promise<any[]> {
        if (!PropertyService._context) {
            throw new Error('PropertyService: no SPFx context for REST calls');
        }
        const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
        const candidates = ['Property', 'Properties', 'property', 'properties'];
        for (const t of candidates) {
            const url = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(t)}')/views`;
            try {
                // eslint-disable-next-line no-console
                console.debug(`PropertyService.fetchViewsViaRest: trying list title='${t}' url=${url}`);
                const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
                // eslint-disable-next-line no-console
                console.debug('PropertyService.fetchViewsViaRest: response status', res.status);
                const data = await res.json();
                // eslint-disable-next-line no-console
                console.debug('PropertyService.fetchViewsViaRest: response body', data);
                const views = data.value || [];

                const detailed = await Promise.all(views.map(async (v: any) => {
                    const viewCopy: any = { ...v };
                    try {
                        const vfUrl = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(t)}')/views/getById('${v.Id}')/ViewFields`;
                        const vfRes = await PropertyService._context!.spHttpClient.get(vfUrl, SPHttpClient.configurations.v1);
                        const vfData = await vfRes.json();
                        viewCopy.ViewFields = vfData.value || (vfData.Items || vfData.Results || vfData.results) || [];
                    } catch (e) {
                        // eslint-disable-next-line no-console
                        console.warn('PropertyService.fetchViewsViaRest: failed to fetch ViewFields for', v.Id, e);
                        viewCopy.ViewFields = [];
                    }
                    return viewCopy;
                }));

                if (detailed && detailed.length > 0) {
                    // eslint-disable-next-line no-console
                    console.debug(`PropertyService.fetchViewsViaRest: found ${detailed.length} views for list '${t}'`);
                    return detailed;
                }
            } catch (e) {
                // eslint-disable-next-line no-console
                console.warn(`PropertyService.fetchViewsViaRest: attempt for '${t}' failed`, e);
            }
        }

        // If none succeeded, try to list lists for debugging
        try {
            const listUrl = `${webUrl}/_api/web/lists?$select=Title,Id,BaseTemplate,Hidden`;
            const listRes = await PropertyService._context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            const listData = await listRes.json();
            // eslint-disable-next-line no-console
            console.debug('PropertyService.fetchViewsViaRest: lists on web', listData.value && listData.value.length ? listData.value.slice(0, 20) : listData.value);
        } catch (e) {
            // eslint-disable-next-line no-console
            console.warn('PropertyService.fetchViewsViaRest: failed to list lists for debugging', e);
        }

        return [];
    }

    private static async fetchItemsViaRest(top = 200): Promise<IPropertyItem[]> {
        if (!PropertyService._context) {
            throw new Error('PropertyService: no SPFx context for REST calls');
        }
        const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
        const url = `${webUrl}/_api/web/lists/getByTitle('Property')/items?$top=${top}`;
        const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await res.json();
        return data.value || [];
    }

    /** Search users on the site for people picker suggestions. Returns minimal user objects. */
    public static async searchUsers(query: string, top = 20): Promise<any[]> {
        if (!PropertyService._context) {
            throw new Error('PropertyService: no SPFx context for REST calls');
        }

        const webUrl = PropertyService._context.pageContext.web.absoluteUrl;
        const q = String(query || '').trim();
        if (!q) return [];

        // Try siteusers with substringof on Title and Email
        // Note: substringof is supported in SharePoint OData endpoint
        const filter = `substringof('${q.replace("'", "''")}',Title) or substringof('${q.replace("'", "''")}',Email)`;
        const url = `${webUrl}/_api/web/siteusers?$filter=${encodeURIComponent(filter)}&$top=${top}&$select=Id,Title,LoginName,Email`;
        try {
            // eslint-disable-next-line no-console
            console.debug('PropertyService.searchUsers: url', url);
            const res = await PropertyService._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            const data = await res.json();
            const users = data && data.value ? data.value : [];
            return users.map((u: any) => ({ id: u.Id, title: u.Title, login: u.LoginName, email: u.Email }));
        } catch (e) {
            // eslint-disable-next-line no-console
            console.warn('PropertyService.searchUsers failed', e);
            return [];
        }
    }
}
