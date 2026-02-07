import * as React from 'react';
import { useEffect, useState } from 'react';
import { IPropertyItem, PropertyService } from '../services/PropertyService';
import {
    PrimaryButton,
    DefaultButton,
    DetailsList,
    IColumn,
    Modal,
    TextField,
    Stack,
    Dropdown,
    IDropdownOption,
    IconButton,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType
    , DatePicker, Checkbox
    , Pivot, PivotItem
} from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import styles from './PropertyManager.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPropertyManagerProps {
    context: WebPartContext;
}

const DEFAULT_VIEWS: { [key: string]: string[] } = {
    Factsheet: ['Title', 'Location', 'UnitCount', 'Manager'],
    Risk: ['Title', 'RiskLevel', 'RiskNotes'],
};

export const PropertyManager: React.FunctionComponent<IPropertyManagerProps> = (props) => {
    const [items, setItems] = useState<IPropertyItem[] | null>(null);
    const [loading, setLoading] = useState<boolean>(false);
    const [error, setError] = useState<string | null>(null);
    const [fields, setFields] = useState<any[] | null>(null);
    const [views, setViews] = useState<any[] | null>(null);

    const [selectedItem, setSelectedItem] = useState<IPropertyItem | null>(null);
    const [isModalOpen, setIsModalOpen] = useState(false);
    const [isEditing, setIsEditing] = useState(false);

    const [selectedView, setSelectedView] = useState<string>('Factsheet');
    const [availableViews, setAvailableViews] = useState<string[]>(Object.keys(DEFAULT_VIEWS));

    // form state keyed by field internal name
    const [formValues, setFormValues] = useState<{ [k: string]: any }>({});
    const [activeTab, setActiveTab] = useState<string>('data');
    const [showRawJson, setShowRawJson] = useState<boolean>(false);

    // people suggestions cache to reduce repeated calls
    const [peopleCache] = useState<Map<string, any[]>>(new Map());
    const [choicesCache, setChoicesCache] = useState<Map<string, IDropdownOption[]>>(new Map());
    const [lookupCache, setLookupCache] = useState<Map<string, IDropdownOption[]>>(new Map());
    const [userCache, setUserCache] = useState<Map<string, string>>(new Map());
    // Bulk update UI state
    const [isBulkModalOpen, setIsBulkModalOpen] = useState(false);
    const [bulkSelectedFields, setBulkSelectedFields] = useState<string[]>([]);
    const [bulkFormValues, setBulkFormValues] = useState<{ [k: string]: any }>({});
    const [bulkLoading, setBulkLoading] = useState(false);
    const [bulkProgress, setBulkProgress] = useState<{ done: number; total: number } | null>(null);

    const resolvePeopleSuggestions = async (filterText: string) => {
        const q = String(filterText || '').trim();
        if (!q) return [];
        if (peopleCache.has(q)) return peopleCache.get(q)!;
        try {
            const users = await PropertyService.searchUsers(q, 20);
            const mapped = users.map(u => ({ key: String(u.id || u.login || u.email || u.title), text: u.title || u.email || u.login, secondaryText: u.email || '', id: u.id }));
            peopleCache.set(q, mapped);
            return mapped;
        } catch (e) {
            return [];
        }
    };

    async function loadSchema() {
        setError(null);
        try {
            const f = await PropertyService.getFields();
            const v = await PropertyService.getViews();
            setFields(f || []);
            const viewsArr = v || [];
            setViews(viewsArr);

            // Build dropdown titles robustly. Use Title if present, otherwise fallback to Id.
            const titles = viewsArr.map((vv: any) => {
                if (!vv) return '';
                // prefer Title, then 'TitleResource' or DisplayName-like props, finally Id
                if (vv.Title) return String(vv.Title);
                if (vv.DisplayName) return String(vv.DisplayName);
                if (vv.TitleResource && vv.TitleResource.Value) return String(vv.TitleResource.Value);
                if (vv.Id) return String(vv.Id);
                return '';
            }).filter((t: string) => t);

            // Always replace the default hardcoded views with actual list views when loadSchema is called
            if (titles.length > 0) {
                setAvailableViews(titles);
                if (titles.indexOf(selectedView) === -1) {
                    setSelectedView(titles[0]);
                }
            } else {
                // No titled views returned — still clear defaults to reflect that we attempted a real fetch
                setAvailableViews(titles.length ? titles : []);
                if ((titles.length === 0) && viewsArr.length > 0) {
                    // fallback: use view Ids so user can at least select them
                    const ids = viewsArr.map((vv: any) => vv && vv.Id ? String(vv.Id) : '').filter((s: string) => s);
                    if (ids.length > 0) {
                        setAvailableViews(ids);
                        setSelectedView(ids[0]);
                    }
                }
            }
            // Debugging aid
            // eslint-disable-next-line no-console
            console.debug('PropertyManager.loadSchema: fetched views', viewsArr);
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Failed to load schema');
        }
    }

    useEffect(() => {
        PropertyService.init(props.context);
        const initAll = async () => {
            // load schema first so default view selection is correct
            await loadSchema();
            await loadItems();
        };
        void initAll();
    }, [props.context]);

    // initialize form values when selectedView or fields change
    useEffect(() => {
        if (!selectedView) return;
        const viewFields = getViewFields(selectedView || undefined);
        const initial: { [k: string]: any } = {};
        viewFields.forEach((f) => { initial[f] = formValues[f] ?? ''; });
        setFormValues(initial);
        // keep tab selection (no inline form shown)
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [selectedView, fields]);

    // When modal opens or schema/selectedView changes, preload choice & lookup options for fields in the view
    useEffect(() => {
        const load = async () => {
            if (!isModalOpen || !fields) return;
            const viewFields = getViewFields(selectedView || undefined);
            for (const f of viewFields) {
                // inline meta resolver (avoid calling helper before it's defined)
                let meta = null as any;
                if (fields && Array.isArray(fields)) {
                    meta = fields.find((ff: any) => ff.InternalName === f || ff.Title === f) || null;
                    if (!meta) {
                        const key = String(f || '').toLowerCase();
                        meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase() === key) || (ff.Title && String(ff.Title).toLowerCase() === key)) || null;
                    }
                    if (!meta) {
                        const key = String(f || '').toLowerCase();
                        meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase().includes(key)) || (ff.Title && String(ff.Title).toLowerCase().includes(key))) || null;
                    }
                }
                if (!meta) continue;
                const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                const type = rawType.toLowerCase();

                // Choice fields
                const isChoice = type.includes('choice') || !!meta.Choices;
                if (isChoice && !choicesCache.has(f)) {
                    try {
                        const fld = await PropertyService.getFieldChoices(f);
                        let choices: any[] = [];
                        if (fld) {
                            if (Array.isArray(fld.Choices)) choices = fld.Choices;
                            if (fld.Choices && Array.isArray(fld.Choices.results)) choices = fld.Choices.results;
                            if (Array.isArray(meta.Choices)) choices = meta.Choices;
                        }
                        const opts: IDropdownOption[] = (choices || []).map((c: any) => ({ key: String(c), text: String(c) }));
                        setChoicesCache(prev => new Map(prev).set(f, opts));
                    } catch (e) {
                        // ignore
                    }
                }

                // Lookup fields
                const isLookup = type.includes('lookup') || (!!meta && !!(meta.LookupList || meta.lookupList)) || (meta && meta.SchemaXml && /LookupList/i.test(String(meta.SchemaXml)));
                if (isLookup && !lookupCache.has(f)) {
                    try {
                        const fld = await PropertyService.getFieldChoices(f);
                        const lookupListId = fld && (fld.LookupList || fld.lookupList || fld.LookupListId || fld.LookupListId) ? (fld.LookupList || fld.lookupList || fld.LookupListId || fld.LookupListId) : null;
                        if (lookupListId) {
                            const items = await PropertyService.getLookupItems(String(lookupListId));
                            const opts: IDropdownOption[] = (items || []).map((it: any) => ({ key: String(it.Id), text: it.Title || it.Title }));
                            setLookupCache(prev => new Map(prev).set(f, opts));
                        }
                    } catch (e) {
                        // ignore
                    }
                }
            }
        };
        void load();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isModalOpen, fields, selectedView]);



    async function loadItems() {
        setLoading(true);
        setError(null);
        try {
            const res = await PropertyService.getItems();
            // Resolve user display names for person fields before setting items
            // so the initial render shows display names instead of numeric ids.
            const nameMap = await resolveUserDisplayNames(res);

            // If we have resolved names, annotate the items so the first render
            // contains display names (don't rely on userCache state being applied
            // before items are rendered).
            if (nameMap && nameMap.size > 0) {
                // find person fields in the schema
                const personFields = (fields || []).filter((f: any) => {
                    const rawType = f && (f.TypeAsString || f.FieldType || f.Type) ? String(f.TypeAsString || f.FieldType || f.Type) : '';
                    const type = rawType.toLowerCase();
                    const fieldTypeKind = f && (f.FieldTypeKind || f.FieldType || f.Type) ? Number(f.FieldTypeKind || f.FieldType || f.Type) : undefined;
                    return fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (f && f.SchemaXml && /UserField/i.test(String(f.SchemaXml)));
                }).map((f: any) => f.InternalName).filter(Boolean) as string[];

                if (personFields.length > 0) {
                    for (const it of res) {
                        for (const pf of personFields) {
                            try {
                                const meta = getMetaForField(pf);
                                const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                                const idVal = (it as any)[`${pf}Id`];
                                const rawVal = (it as any)[pf];

                                if (allowMulti) {
                                    let ids: any[] = [];
                                    if (Array.isArray(idVal) && idVal.length > 0) ids = idVal;
                                    else if (idVal && Array.isArray(idVal.results)) ids = idVal.results;
                                    else if (typeof idVal === 'string' && idVal.trim()) ids = idVal.split(',').map((s: string) => s.trim()).filter((s: string) => s);
                                    else if (typeof idVal === 'number') ids = [idVal];

                                    if (ids.length > 0 && (!Array.isArray(rawVal) || rawVal.length === 0)) {
                                        // replace rawVal with array of objects containing Title so renderer shows names
                                        (it as any)[pf] = ids.map((id: any) => ({ Title: nameMap.get(String(id)) ?? String(id), Id: Number(id) }));
                                    }
                                } else {
                                    // single user
                                    if ((!rawVal || Object.keys(rawVal).length === 0) && (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim()))) {
                                        const id = Array.isArray(idVal) && idVal.length > 0 ? idVal[0] : idVal;
                                        (it as any)[pf] = { Title: nameMap.get(String(id)) ?? String(id), Id: Number(id) };
                                    }
                                }
                            } catch (err) {
                                // ignore per-item annotation errors
                            }
                        }
                    }
                }
            }

            setItems(res);
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Failed to load items');
        } finally {
            setLoading(false);
        }
    }

    async function resolveUserDisplayNames(items: IPropertyItem[] | null): Promise<Map<string, string> | null> {
        if (!items || items.length === 0 || !fields) return null;
        try {
            // find person fields in the schema
            const personFields = (fields || []).filter((f: any) => {
                const rawType = f && (f.TypeAsString || f.FieldType || f.Type) ? String(f.TypeAsString || f.FieldType || f.Type) : '';
                const type = rawType.toLowerCase();
                const fieldTypeKind = f && (f.FieldTypeKind || f.FieldType || f.Type) ? Number(f.FieldTypeKind || f.FieldType || f.Type) : undefined;
                return fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (f && f.SchemaXml && /UserField/i.test(String(f.SchemaXml)));
            }).map((f: any) => f.InternalName).filter(Boolean) as string[];

            if (personFields.length === 0) return null;

            const idsToResolve = new Set<string>();
            for (const it of items) {
                for (const pf of personFields) {
                    const idVal = (it as any)[`${pf}Id`];
                    const rawVal = (it as any)[pf];
                    if (Array.isArray(idVal) && idVal.length > 0) {
                        idVal.forEach((id: any) => idsToResolve.add(String(id)));
                    } else if (idVal && Array.isArray(idVal.results)) {
                        idVal.results.forEach((id: any) => idsToResolve.add(String(id)));
                    } else if (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim())) {
                        idsToResolve.add(String(idVal));
                    } else if (rawVal && (rawVal.Id || rawVal.ID || rawVal.id)) {
                        idsToResolve.add(String(rawVal.Id ?? rawVal.ID ?? rawVal.id));
                    }
                }
            }

            if (idsToResolve.size === 0) return null;

            const missing: string[] = [];
            const current = new Map(userCache);
            idsToResolve.forEach(id => { if (!current.has(id)) missing.push(id); });
            if (missing.length === 0) return null;

            const resolved = await Promise.all(missing.map(async (id) => {
                try {
                    const u = await PropertyService.getUserById(Number(id));
                    return { id: String(id), name: u && (u.title || u.Title || u.login || u.email) ? (u.title || u.Title || u.login || u.email) : String(id) };
                } catch (e) {
                    return { id: String(id), name: String(id) };
                }
            }));

            resolved.forEach(r => current.set(r.id, r.name));
            setUserCache(current);
            return current;
        } catch (e) {
            // ignore resolution errors
            return null;
        }
    }

    // (removed) resolveMissingIds: on-demand id resolver was removed due to stability issues

    const exportViewToCsv = async () => {
        if (!items || items.length === 0) {
            setError('No items available to export');
            return;
        }
        try {
            const viewFields = getViewFields(selectedView || undefined);

            // Ensure any missing user names are resolved before building CSV
            const idsToResolve = new Set<string>();
            for (const it of items) {
                for (const f of viewFields) {
                    const meta = getMetaForField(f);
                    const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                    const type = rawType.toLowerCase();
                    const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
                    const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                    const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.TypeAsString && /user/i.test(String(meta.TypeAsString))) || (meta && typeof meta.SchemaXml === 'string' && /UserField/i.test(meta.SchemaXml));
                    if (!isPersonField) continue;
                    const idVal = (it as any)[`${f}Id`];
                    if (!idVal) continue;
                    if (Array.isArray(idVal) && idVal.length > 0) idVal.forEach((x: any) => idsToResolve.add(String(x)));
                    else if (idVal && Array.isArray(idVal.results)) idVal.results.forEach((x: any) => idsToResolve.add(String(x)));
                    else if (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim())) idsToResolve.add(String(idVal));
                }
            }
            if (idsToResolve.size > 0) await resolveUserDisplayNames(items);

            const headerLabels = viewFields.map(f => (getMetaForField(f)?.Title) || f);

            const escapeCsv = (v: any) => {
                if (v === null || typeof v === 'undefined') return '';
                const s = String(v).replace(/"/g, '""');
                return `"${s}"`;
            };

            const rows = items.map(it => {
                const row: any = {};
                for (const f of viewFields) {
                    const meta = getMetaForField(f);
                    const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                    const type = rawType.toLowerCase();
                    const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
                    const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                    const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.TypeAsString && /user/i.test(String(meta.TypeAsString))) || (meta && typeof meta.SchemaXml === 'string' && /UserField/i.test(meta.SchemaXml));

                    let out = '';
                    if (isPersonField) {
                        const rawVal = (it as any)[f];
                        const idVal = (it as any)[`${f}Id`];
                        if (allowMulti) {
                            let names: string[] = [];
                            if (Array.isArray(rawVal) && rawVal.length > 0) {
                                names = rawVal.map((u: any) => u && (u.Title || u.title || u.Name || u.LoginName || u.Email) ? (u.Title || u.title || u.Name || u.LoginName || u.Email) : String(u && (u.Id || u.ID || u.id) || '')).filter(Boolean);
                            } else if (idVal && Array.isArray(idVal.results)) {
                                names = idVal.results.map((id: any) => userCache && userCache.has(String(id)) ? userCache.get(String(id))! : String(id));
                            } else if (Array.isArray(idVal)) {
                                names = idVal.map((id: any) => userCache && userCache.has(String(id)) ? userCache.get(String(id))! : String(id));
                            } else if (typeof idVal === 'string' && idVal.indexOf(',') >= 0) {
                                names = idVal.split(',').map((s: string) => s.trim()).filter((s: string) => s).map(s => userCache && userCache.has(s) ? userCache.get(s)! : s);
                            } else if (typeof idVal === 'string' && idVal.trim()) {
                                const s = idVal.trim();
                                names = [userCache && userCache.has(s) ? userCache.get(s)! : s];
                            }
                            out = names.join('; ');
                        } else {
                            if (rawVal && (rawVal.Title || rawVal.title || rawVal.Name || rawVal.LoginName || rawVal.Email)) {
                                out = rawVal.Title || rawVal.title || rawVal.Name || rawVal.LoginName || rawVal.Email || '';
                            } else if (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim())) {
                                const s = String(idVal);
                                out = (userCache && userCache.has(s)) ? userCache.get(s)! : s;
                            } else if (typeof rawVal === 'string' && rawVal.trim()) {
                                out = rawVal;
                            }
                        }
                    } else {
                        const v = (it as any)[f];
                        if (Array.isArray(v)) out = v.join('; ');
                        else if (v && typeof v === 'object') out = v.Title || v.Title || JSON.stringify(v);
                        else out = (v === null || typeof v === 'undefined') ? '' : String(v);
                    }
                    row[f] = out;
                }
                return row;
            });

            const csvLines: string[] = [];
            csvLines.push(headerLabels.map(h => escapeCsv(h)).join(','));
            for (const r of rows) {
                const line = viewFields.map(f => escapeCsv(r[f])).join(',');
                csvLines.push(line);
            }

            const csv = csvLines.join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const nameSafe = String(selectedView || 'export').replace(/[^a-z0-9_-]/gi, '_');
            a.download = `${nameSafe}.csv`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Export failed');
        }
    };

    const openCreate = async () => {
        setSelectedItem(null);
        // ensure schema/fields are loaded so field metadata (person types) are detected
        if (!fields || fields.length === 0) {
            // loadSchema sets `fields` state
            // fire and await so we have metadata before building the form
            // ignore errors here: loadSchema will set error state if it fails
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            await loadSchema();
        }
        // initialize form values for all fields in the currently selected view
        const viewFields = getViewFields(selectedView || undefined);
        const initial: { [k: string]: any } = {};
        viewFields.forEach((f) => { initial[f] = ''; });
        setFormValues(initial);
        setIsEditing(false);
        setIsModalOpen(true);
    };

    const openEdit = async (item: IPropertyItem) => {
        setSelectedItem(item);
        if (!fields || fields.length === 0) {
            // ensure field metadata is present before rendering edit form
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            await loadSchema();
        }

        // Prepare form values and map person fields into PeoplePicker-friendly shapes
        const viewFields = getViewFields(selectedView || undefined);
        const initialValues: { [k: string]: any } = { ...item };
        if (fields && Array.isArray(fields)) {
            for (const f of viewFields) {
                const meta = getMetaForField(f);
                if (!meta) continue;
                const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                const type = rawType.toLowerCase();
                const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
                const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.SchemaXml && /UserField/i.test(String(meta.SchemaXml)));
                if (!isPersonField) continue;

                const rawVal = (item as any)[f];
                const idVal = (item as any)[`${f}Id`];

                const toPickerItem = (u: any) => {
                    if (!u) return null;
                    const id = u.Id ?? u.ID ?? u.id ?? u.Key ?? u.key ?? undefined;
                    const title = u.Title ?? u.title ?? u.Name ?? u.LoginName ?? '';
                    const email = u.Email ?? u.email ?? '';
                    const display = title || email || String(id ?? '');
                    return { id: id ?? display, key: id ?? display, text: display, secondaryText: email || '' };
                };

                if (allowMulti) {
                    let mapped: any[] = [];
                    if (Array.isArray(rawVal) && rawVal.length > 0) {
                        mapped = rawVal.map((u: any) => toPickerItem(u)).filter(Boolean);
                    } else {
                        // Extract ids from various shapes
                        let ids: any[] = [];
                        if (idVal && Array.isArray(idVal.results)) ids = idVal.results;
                        else if (idVal && Array.isArray(idVal)) ids = idVal;
                        else if (typeof idVal === 'string' && idVal.trim()) ids = idVal.split(',').map((s: string) => s.trim()).filter((s: string) => s);
                        else if (typeof idVal === 'number') ids = [idVal];

                        if (ids.length > 0) {
                            const resolved = await Promise.all(ids.map(async (id) => {
                                try {
                                    const u = await PropertyService.getUserById(Number(id));
                                    const pi = toPickerItem(u);
                                    return pi ?? { id, key: String(id), text: String(id), secondaryText: '' };
                                } catch (e) {
                                    return { id, key: String(id), text: String(id), secondaryText: '' };
                                }
                            }));
                            mapped = resolved.filter(Boolean);
                        }
                    }
                    initialValues[f] = mapped;
                } else {
                    // single user
                    if (rawVal && (rawVal.Id || rawVal.ID || rawVal.Title || rawVal.Title === '')) {
                        const mapped = toPickerItem(rawVal);
                        initialValues[f] = mapped ? mapped : rawVal;
                    } else if (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim())) {
                        // we only have the id; try to resolve it to a user object so the picker shows a display name
                        const id = (Array.isArray(idVal) && idVal.length > 0) ? idVal[0] : idVal;
                        try {
                            const u = await PropertyService.getUserById(Number(id));
                            const mapped = toPickerItem(u);
                            if (mapped) initialValues[f] = mapped;
                            else initialValues[f] = { id, key: String(id), text: String(id), secondaryText: '' };
                        } catch (e) {
                            initialValues[f] = { id, key: String(id), text: String(id), secondaryText: '' };
                        }
                    } else if (typeof rawVal === 'string' && rawVal.trim()) {
                        // sometimes the value is stored as a display string
                        initialValues[f] = { id: rawVal, key: rawVal, text: rawVal, secondaryText: '' };
                    }
                }
            }
        }

        setFormValues(initialValues);
        setIsEditing(true);
        setIsModalOpen(true);
    };

    const doDelete = async (item: IPropertyItem) => {
        if (!item || !item.Id) return;
        if (!confirm(`Delete property ${item.Title || item.Id}?`)) return;
        try {
            setLoading(true);
            await PropertyService.deleteItem(item.Id as number);
            await loadItems();
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Delete failed');
        } finally {
            setLoading(false);
        }
    };
    // Helper shared with `save` and bulk updates: prepare a payload containing only saveable fields (person fields mapped to Ids)
    const prepareDataForSave = (values: { [k: string]: any }) => {
        const data: { [k: string]: any } = { ...values };
        if (fields && Array.isArray(fields)) {
            Object.keys(values).forEach((k) => {
                const meta = getMetaForField(k);
                if (!meta) return;
                const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                const type = rawType.toLowerCase();
                const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
                const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.SchemaXml && /UserField/i.test(String(meta.SchemaXml)));

                if (isPersonField) {
                    const val = values[k];
                    if (allowMulti) {
                        const ids = Array.isArray(val) ? val.map((v: any) => Number(v?.id ?? v?.key ?? v)).filter((n: any) => !Number.isNaN(n)) : [];
                        data[`${k}Id`] = { results: ids };
                    } else {
                        const id = Array.isArray(val) ? Number(val[0]?.id ?? val[0]?.key ?? val[0]) : Number(val?.id ?? val?.key ?? val);
                        if (!Number.isNaN(id)) data[`${k}Id`] = id;
                    }
                    // remove the display value to avoid conflicts when saving
                    delete data[k];
                }
            });
        }
        return data;
    };

    const save = async () => {
        setLoading(true);
        setError(null);
        try {
            // prepareDataForSave is now a top-level helper (see above)

            const payload = prepareDataForSave(formValues);
            if (isEditing && selectedItem && selectedItem.Id) {
                await PropertyService.updateItem(selectedItem.Id as number, payload);
            } else {
                await PropertyService.createItem(payload);
            }
            setIsModalOpen(false);
            await loadItems();
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Save failed');
        } finally {
            setLoading(false);
        }
    };



    function getViewFields(v?: string): string[] {
        if (!v) return ['Title'];
        // Prefer fields from fetched views (if available)
        if (views && views.length > 0) {
            // Try to find by Title first, then by Id
            let found = (views as any[]).find(x => x && x.Title === v);
            if (!found) {
                found = (views as any[]).find(x => x && (String(x.Id) === String(v) || x.DisplayName === v));
            }
            if (found) {
                const vf = found.ViewFields;
                // ViewFields sometimes comes as an array, or as an object with Items, or as a string
                if (Array.isArray(vf) && vf.length > 0) return vf as string[];
                if (vf && Array.isArray(vf.Items) && vf.Items.length > 0) return vf.Items as string[];
                // Some responses use 'Results' or 'results' or comma-separated strings
                if (vf && Array.isArray(vf.Results) && vf.Results.length > 0) return vf.Results as string[];
                if (vf && Array.isArray(vf.results) && vf.results.length > 0) return vf.results as string[];
                if (typeof vf === 'string' && vf.trim().length > 0) {
                    return vf.split(',').map((s: string) => s.trim()).filter((s: string) => s);
                }
            }
        }

        return (DEFAULT_VIEWS as any)[v] || ['Title'];
    }

    // helper to find field metadata for a given view field name (same logic as renderFieldControl uses)
    function getMetaForField(fieldInternalName: string) {
        let meta = null as any;
        if (fields && Array.isArray(fields)) {
            meta = fields.find((ff: any) => ff.InternalName === fieldInternalName || ff.Title === fieldInternalName) || null;
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase() === key) || (ff.Title && String(ff.Title).toLowerCase() === key)) || null;
            }
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase().includes(key)) || (ff.Title && String(ff.Title).toLowerCase().includes(key))) || null;
            }
        }
        return meta;
    }

    const columnsForView = (): IColumn[] => {
        const viewCols = getViewFields(selectedView);
        const cols: IColumn[] = viewCols.map((k, i) => ({ key: k, name: k, fieldName: k, minWidth: 100, maxWidth: 300, isResizable: true }));
        // Add a single actions column (one set of icons per row)
        cols.push({
            key: 'actions',
            name: '',
            fieldName: 'actions',
            minWidth: 60,
            maxWidth: 120,
            isResizable: false,
            onRender: (item?: any) => {
                if (!item) return null;
                return (
                    <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => openEdit(item)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={() => doDelete(item)} />
                    </div>
                );
            }
        });

        return cols;
    };

    const onFormFieldChange = (field: string, value: any) => {
        setFormValues(prev => ({ ...prev, [field]: value }));
    };
    function renderFieldControl(fieldInternalName: string) {
        // find field metadata if available. Try exact matches, then case-insensitive, then contains.
        let meta = null as any;
        if (fields && Array.isArray(fields)) {
            meta = fields.find((ff: any) => ff.InternalName === fieldInternalName || ff.Title === fieldInternalName) || null;
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase() === key) || (ff.Title && String(ff.Title).toLowerCase() === key)) || null;
            }
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase().includes(key)) || (ff.Title && String(ff.Title).toLowerCase().includes(key))) || null;
            }
        }
        const value = formValues[fieldInternalName] ?? '';
        // debug: show metadata for this field to aid detection (logs after type is computed)

        // normalize type string for flexible matching
        const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
        const type = rawType.toLowerCase();
        // debug: show metadata for this field to aid detection
        // eslint-disable-next-line no-console
        console.debug('renderFieldControl:', fieldInternalName, { meta, type, value });

        // Boolean
        if (type === 'boolean' || type === 'bool') {
            return (
                <Checkbox
                    checked={!!value}
                    label={meta?.Title || fieldInternalName}
                    onChange={(_, checked) => onFormFieldChange(fieldInternalName, !!checked)}
                />
            );
        }

        // Date/time — accept several possible type names
        if (type.includes('date') || type.includes('datetime')) {
            return (
                <DatePicker
                    placeholder={meta?.Title || fieldInternalName}
                    value={value ? new Date(value) : undefined}
                    onSelectDate={(d) => onFormFieldChange(fieldInternalName, d ? d.toISOString() : '')}
                    styles={{ root: { width: '100%' } }}
                />
            );
        }

        // Number
        if (type.includes('number') || type.includes('int') || type.includes('float') || type.includes('decimal')) {
            return (
                <TextField
                    type="number"
                    label={meta?.Title || fieldInternalName}
                    value={String(value ?? '')}
                    onChange={(_, v) => onFormFieldChange(fieldInternalName, v ? Number(v) : undefined)}
                    styles={{ root: { width: '100%' } }}
                />
            );
        }

        // Multi-line text / Note
        if (type.includes('note') || type.includes('multiline') || (meta && (meta.MultipleLines || meta.AllowMultipleLines))) {
            return (
                <TextField
                    label={meta?.Title || fieldInternalName}
                    multiline
                    rows={4}
                    value={String(value ?? '')}
                    onChange={(_, v) => onFormFieldChange(fieldInternalName, v)}
                    styles={{ root: { width: '100%' } }}
                />
            );
        }

        // Person / User fields (single or multi)
        const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
        const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
        const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.TypeAsString && /user/i.test(String(meta.TypeAsString))) || (meta && typeof meta.SchemaXml === 'string' && /UserField/i.test(meta.SchemaXml));
        // eslint-disable-next-line no-console
        if (isPersonField) console.debug('renderFieldControl: detected person field', fieldInternalName, { fieldTypeKind, allowMulti, meta });
        if (isPersonField) {
            // Use NormalPeoplePicker from Fluent UI if available — fall back to a simple TextField when suggestions are not provided.
            try {
                return (
                    <div>
                        <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                        <NormalPeoplePicker
                            onResolveSuggestions={async (filterText) => await resolvePeopleSuggestions(filterText)}
                            getTextFromItem={(item: any) => item && item.text ? item.text : ''}
                            pickerSuggestionsProps={{ suggestionsHeaderText: 'People', noResultsFoundText: 'No matches' }}
                            selectedItems={Array.isArray(value) ? value.map((v: any) => (typeof v === 'string' ? { key: v, text: v } : (v && v.text ? { key: v.id || v.key || v.text, text: v.text, secondaryText: v.secondaryText } : { key: String(v), text: String(v) }))) : (value ? (typeof value === 'string' ? [{ key: value, text: value }] : [{ key: value.id || value.key || String(value), text: value.text || String(value) }]) : [])}
                            onChange={(items: any) => onFormFieldChange(fieldInternalName, items ? items.map((i: any) => ({ id: i.key, text: i.text, secondaryText: i.secondaryText })) : [])}
                            resolveDelay={300}
                            styles={{ root: { minWidth: '100%' } }}
                        />
                    </div>
                );
            } catch (err) {
                // If NormalPeoplePicker isn't available or fails, fallback to a text input that accepts display names/emails
                return (
                    <TextField
                        label={meta?.Title || fieldInternalName}
                        placeholder="Enter user (display name or email)"
                        value={Array.isArray(value) ? (value as any[]).join('; ') : String(value ?? '')}
                        onChange={(_, v) => onFormFieldChange(fieldInternalName, v)}
                    />
                );
            }
        }

        // Choice fields (single or multi-select)
        const isChoiceField = (type.includes('choice') || (meta && (meta.Choices || meta.Choices && meta.Choices.results)));
        if (isChoiceField) {
            const opts = choicesCache.get(fieldInternalName) || ([] as IDropdownOption[]);
            const multi = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
            // compute selected value(s)
            const selected = value ?? (multi ? [] : '');
            return (
                <div>
                    <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                    <Dropdown
                        options={opts}
                        selectedKeys={multi ? (Array.isArray(selected) ? selected : []) : undefined}
                        selectedKey={!multi ? (selected ? String(selected) : undefined) : undefined}
                        multiSelect={multi}
                        placeholder={meta?.Title || fieldInternalName}
                        onChange={(_, option) => {
                            if (!option) return;
                            if (multi) {
                                const cur = Array.isArray(formValues[fieldInternalName]) ? [...formValues[fieldInternalName]] : [];
                                const key = option.key as string;
                                if (option.selected) {
                                    if (cur.indexOf(key) === -1) cur.push(key);
                                } else {
                                    const idx = cur.indexOf(key);
                                    if (idx >= 0) cur.splice(idx, 1);
                                }
                                onFormFieldChange(fieldInternalName, cur);
                            } else {
                                onFormFieldChange(fieldInternalName, option.key);
                            }
                        }}
                        styles={{ root: { minWidth: '100%' } }}
                    />
                </div>
            );
        }

        // Lookup fields
        const isLookupField = (type.includes('lookup') || (!!meta && !!(meta.LookupList || meta.lookupList)) || (meta && meta.SchemaXml && /LookupList/i.test(String(meta.SchemaXml))));
        if (isLookupField) {
            const opts = lookupCache.get(fieldInternalName) || ([] as IDropdownOption[]);
            const multi = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
            const selected = value ?? (multi ? [] : '');
            return (
                <div>
                    <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                    <Dropdown
                        options={opts}
                        selectedKeys={multi ? (Array.isArray(selected) ? selected : []) : undefined}
                        selectedKey={!multi ? (selected ? String(selected) : undefined) : undefined}
                        multiSelect={multi}
                        placeholder={meta?.Title || fieldInternalName}
                        onChange={(_, option) => {
                            if (!option) return;
                            if (multi) {
                                const cur = Array.isArray(formValues[fieldInternalName]) ? [...formValues[fieldInternalName]] : [];
                                const key = option.key as string;
                                if (option.selected) {
                                    if (cur.indexOf(key) === -1) cur.push(key);
                                } else {
                                    const idx = cur.indexOf(key);
                                    if (idx >= 0) cur.splice(idx, 1);
                                }
                                onFormFieldChange(fieldInternalName, cur);
                            } else {
                                onFormFieldChange(fieldInternalName, option.key);
                            }
                        }}
                        styles={{ root: { minWidth: '100%' } }}
                    />
                </div>
            );
        }

        // default: text field
        return (
            <TextField
                label={meta?.Title || fieldInternalName}
                value={String(value ?? '')}
                onChange={(_, v) => onFormFieldChange(fieldInternalName, v)}
            />
        );
    }

    // Render control for bulk-update modal (operates on bulkFormValues)
    function renderBulkFieldControl(fieldInternalName: string) {
        let meta = null as any;
        if (fields && Array.isArray(fields)) {
            meta = fields.find((ff: any) => ff.InternalName === fieldInternalName || ff.Title === fieldInternalName) || null;
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase() === key) || (ff.Title && String(ff.Title).toLowerCase() === key)) || null;
            }
            if (!meta) {
                const key = String(fieldInternalName || '').toLowerCase();
                meta = fields.find((ff: any) => (ff.InternalName && String(ff.InternalName).toLowerCase().includes(key)) || (ff.Title && String(ff.Title).toLowerCase().includes(key))) || null;
            }
        }
        const value = bulkFormValues[fieldInternalName] ?? '';
        const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
        const type = rawType.toLowerCase();

        // Boolean
        if (type === 'boolean' || type === 'bool') {
            return (
                <Checkbox
                    checked={!!value}
                    label={meta?.Title || fieldInternalName}
                    onChange={(_, checked) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: !!checked }))}
                />
            );
        }

        // Date/time
        if (type.includes('date') || type.includes('datetime')) {
            return (
                <DatePicker
                    placeholder={meta?.Title || fieldInternalName}
                    value={value ? new Date(value) : undefined}
                    onSelectDate={(d) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: d ? d.toISOString() : '' }))}
                    styles={{ root: { width: '100%' } }}
                />
            );
        }

        // Multi-line text / Note
        if (type.includes('note') || type.includes('multiline') || (meta && (meta.MultipleLines || meta.AllowMultipleLines))) {
            return (
                <TextField
                    label={meta?.Title || fieldInternalName}
                    multiline
                    rows={4}
                    value={String(value ?? '')}
                    onChange={(_, v) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: v }))}
                    styles={{ root: { width: '100%' } }}
                />
            );
        }

        // Person fields (simple support: allow entering a single id or use people picker)
        const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
        const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.TypeAsString && /user/i.test(String(meta.TypeAsString))) || (meta && typeof meta.SchemaXml === 'string' && /UserField/i.test(meta.SchemaXml));
        if (isPersonField) {
            try {
                return (
                    <div>
                        <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                        <NormalPeoplePicker
                            onResolveSuggestions={async (filterText) => await resolvePeopleSuggestions(filterText)}
                            getTextFromItem={(item: any) => item && item.text ? item.text : ''}
                            pickerSuggestionsProps={{ suggestionsHeaderText: 'People', noResultsFoundText: 'No matches' }}
                            selectedItems={Array.isArray(value) ? value.map((v: any) => (typeof v === 'string' ? { key: v, text: v } : (v && v.text ? { key: v.id || v.key || v.text, text: v.text } : { key: String(v), text: String(v) }))) : (value ? (typeof value === 'string' ? [{ key: value, text: value }] : [{ key: value.id || value.key || String(value), text: value.text || String(value) }]) : [])}
                            onChange={(items: any) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: items ? items.map((i: any) => ({ id: i.key, text: i.text, secondaryText: i.secondaryText })) : [] }))}
                            resolveDelay={300}
                            styles={{ root: { minWidth: '100%' } }}
                        />
                    </div>
                );
            } catch (err) {
                return (
                    <TextField
                        label={meta?.Title || fieldInternalName}
                        placeholder="Enter user (display name or email)"
                        value={Array.isArray(value) ? (value as any[]).join('; ') : String(value ?? '')}
                        onChange={(_, v) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: v }))}
                    />
                );
            }
        }

        // Choice and lookup fields: show dropdown if options are cached
        const isChoiceField = (type.includes('choice') || (meta && (meta.Choices || meta.Choices && meta.Choices.results)));
        if (isChoiceField) {
            const opts = choicesCache.get(fieldInternalName) || ([] as IDropdownOption[]);
            const multi = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
            const selected = value ?? (multi ? [] : '');
            return (
                <div>
                    <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                    <Dropdown
                        options={opts}
                        selectedKeys={multi ? (Array.isArray(selected) ? selected : []) : undefined}
                        selectedKey={!multi ? (selected ? String(selected) : undefined) : undefined}
                        multiSelect={multi}
                        placeholder={meta?.Title || fieldInternalName}
                        onChange={(_, option) => {
                            if (!option) return;
                            if (multi) {
                                const cur = Array.isArray(bulkFormValues[fieldInternalName]) ? [...bulkFormValues[fieldInternalName]] : [];
                                const key = option.key as string;
                                if (option.selected) {
                                    if (cur.indexOf(key) === -1) cur.push(key);
                                } else {
                                    const idx = cur.indexOf(key);
                                    if (idx >= 0) cur.splice(idx, 1);
                                }
                                setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: cur }));
                            } else {
                                setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: option.key }));
                            }
                        }}
                        styles={{ root: { minWidth: '100%' } }}
                    />
                </div>
            );
        }

        // Lookup fields
        const isLookupField = (type.includes('lookup') || (!!meta && !!(meta.LookupList || meta.lookupList)) || (meta && meta.SchemaXml && /LookupList/i.test(String(meta.SchemaXml))));
        if (isLookupField) {
            const opts = lookupCache.get(fieldInternalName) || ([] as IDropdownOption[]);
            const multi = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
            const selected = value ?? (multi ? [] : '');
            return (
                <div>
                    <div style={{ marginBottom: 6, fontSize: 12, color: '#333' }}>{meta?.Title || fieldInternalName}</div>
                    <Dropdown
                        options={opts}
                        selectedKeys={multi ? (Array.isArray(selected) ? selected : []) : undefined}
                        selectedKey={!multi ? (selected ? String(selected) : undefined) : undefined}
                        multiSelect={multi}
                        placeholder={meta?.Title || fieldInternalName}
                        onChange={(_, option) => {
                            if (!option) return;
                            if (multi) {
                                const cur = Array.isArray(bulkFormValues[fieldInternalName]) ? [...bulkFormValues[fieldInternalName]] : [];
                                const key = option.key as string;
                                if (option.selected) {
                                    if (cur.indexOf(key) === -1) cur.push(key);
                                } else {
                                    const idx = cur.indexOf(key);
                                    if (idx >= 0) cur.splice(idx, 1);
                                }
                                setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: cur }));
                            } else {
                                setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: option.key }));
                            }
                        }}
                        styles={{ root: { minWidth: '100%' } }}
                    />
                </div>
            );
        }

        // default: text field
        return (
            <TextField
                label={meta?.Title || fieldInternalName}
                value={String(value ?? '')}
                onChange={(_, v) => setBulkFormValues(prev => ({ ...prev, [fieldInternalName]: v }))}
            />
        );
    }

    // Perform bulk update: apply bulkFormValues to all items for bulkSelectedFields
    const performBulkUpdate = async () => {
        if (!items || items.length === 0) {
            setBulkLoading(false);
            setError('No items available to update');
            return;
        }
        if (!bulkSelectedFields || bulkSelectedFields.length === 0) {
            setError('No fields selected for bulk update');
            return;
        }

        setBulkLoading(true);
        setBulkProgress({ done: 0, total: items.length });
        try {
            let done = 0;
            for (const it of items) {
                // build values object only for selected fields
                const values: { [k: string]: any } = {};
                for (const f of bulkSelectedFields) {
                    if (typeof bulkFormValues[f] !== 'undefined') values[f] = bulkFormValues[f];
                }
                console.log('Bulk update - values to prepare:', values, 'bulkFormValues:', bulkFormValues);
                const payload = prepareDataForSave(values);
                console.log('Bulk update - prepared payload:', payload);
                // only call update if payload contains keys
                if (payload && Object.keys(payload).length > 0) {
                    try {
                        console.log('Calling updateItem for item', it.Id, 'with payload:', payload);
                        await PropertyService.updateItem((it as any).Id as number, payload);
                        console.log('Successfully updated item', it.Id);
                    } catch (e) {
                        // ignore per-item failures but surface a generic error later
                        // continue processing remaining items
                        console.error('Error updating item', it.Id, ':', e);
                    }
                } else {
                    console.log('Skipping item', it.Id, '- empty payload');
                }
                done += 1;
                setBulkProgress({ done, total: items.length });
            }

            // refresh list after bulk update
            await loadItems();
            setIsBulkModalOpen(false);
        } catch (e: any) {
            setError(e && e.message ? e.message : 'Bulk update failed');
        } finally {
            setBulkLoading(false);
            setBulkProgress(null);
        }
    };

    return (
        <div className={styles.propertyManager}>
            <div className={styles.pmHeader}>
                <div>
                    <h2 style={{ margin: 0 }}>Property Manager</h2>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                    <PrimaryButton text="Export CSV" onClick={exportViewToCsv} />
                    <PrimaryButton text="Bulk Update" onClick={() => { setBulkSelectedFields([]); setBulkFormValues({}); setIsBulkModalOpen(true); }} />
                    <PrimaryButton text="New" onClick={openCreate} />
                    <DefaultButton text="Refresh" onClick={loadItems} />
                    <DefaultButton text="Reload Schema" onClick={loadSchema} />
                </div>
            </div>

            {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>{error}</MessageBar>}

            <div className={styles.pmLayout}>
                <div className={styles.leftColumn}>
                    <div style={{ marginBottom: 12 }}>
                        <PrimaryButton text="New Item" onClick={openCreate} />
                    </div>
                    <div className={styles.viewsList}>
                        {availableViews.map(v => (
                            <DefaultButton
                                key={v}
                                className={`${styles.viewButton} ${selectedView === v ? styles.viewButtonActive : ''}`}
                                text={v}
                                onClick={() => setSelectedView(v)}
                            />
                        ))}
                    </div>
                </div>

                <div className={styles.bodyColumn}>
                    <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <h3 style={{ margin: 0 }}>{selectedView || 'View'}</h3>
                        <div>
                            <PrimaryButton text="Export CSV" onClick={exportViewToCsv} />
                        </div>
                    </div>

                    <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey || 'data')}>
                        <PivotItem headerText="Data" itemKey="data">
                            <div style={{ marginTop: 6 }}>
                                {loading && <Spinner size={SpinnerSize.medium} label="Loading..." />}
                                {!loading && items && (
                                    <DetailsList
                                        items={items}
                                        columns={columnsForView()}
                                        setKey="set"
                                        selectionMode={0}
                                        onItemInvoked={(item) => openEdit(item as IPropertyItem)}
                                        styles={{ root: { marginTop: 8 } }}
                                        onRenderItemColumn={(item?: any, index?: number, column?: IColumn) => {
                                            if (!item || !column) return null;
                                            const fieldName = (column.fieldName || column.key || '');
                                            // If this is the dedicated actions column, let the column's onRender handle it
                                            if (String(fieldName || '').toLowerCase() === 'actions') return null;
                                            // Try to render person fields (single/multi) with display names
                                            const meta = getMetaForField(fieldName);
                                            const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                                            const type = rawType.toLowerCase();
                                            const fieldTypeKind = meta && (meta.FieldTypeKind || meta.FieldType || meta.Type) ? Number(meta.FieldTypeKind || meta.FieldType || meta.Type) : undefined;
                                            const allowMulti = !!(meta && (meta.AllowMultipleValues || meta.AllowMultiple));
                                            const isPersonField = fieldTypeKind === 20 || type.includes('user') || type.includes('person') || (meta && meta.TypeAsString && /user/i.test(String(meta.TypeAsString))) || (meta && typeof meta.SchemaXml === 'string' && /UserField/i.test(meta.SchemaXml));

                                            let displayValue = '';
                                            if (isPersonField) {
                                                const rawVal = item[fieldName];
                                                const idVal = item[`${fieldName}Id`];
                                                if (allowMulti) {
                                                    let names: string[] = [];
                                                    if (Array.isArray(rawVal) && rawVal.length > 0) {
                                                        names = rawVal.map((u: any) => u && (u.Title || u.title || u.Name || u.LoginName || u.Email) ? (u.Title || u.title || u.Name || u.LoginName || u.Email) : String(u && (u.Id || u.ID || u.id) || '')).filter(Boolean);
                                                    } else if (idVal && Array.isArray(idVal.results)) {
                                                        names = idVal.results.map((id: any) => userCache && userCache.has(String(id)) ? userCache.get(String(id))! : String(id));
                                                    } else if (Array.isArray(idVal)) {
                                                        names = idVal.map((id: any) => userCache && userCache.has(String(id)) ? userCache.get(String(id))! : String(id));
                                                    } else if (typeof idVal === 'string' && idVal.indexOf(',') >= 0) {
                                                        names = idVal.split(',').map((s: string) => s.trim()).filter((s: string) => s).map(s => userCache && userCache.has(s) ? userCache.get(s)! : s);
                                                    } else if (typeof idVal === 'string' && idVal.trim()) {
                                                        const s = idVal.trim();
                                                        names = [(userCache && userCache.has(s) ? userCache.get(s)! : s)];
                                                    }
                                                    displayValue = names.join(', ');
                                                } else {
                                                    if (rawVal && (rawVal.Title || rawVal.title || rawVal.Name || rawVal.LoginName || rawVal.Email)) {
                                                        displayValue = rawVal.Title || rawVal.title || rawVal.Name || rawVal.LoginName || rawVal.Email || '';
                                                    } else if (typeof idVal === 'number' || (typeof idVal === 'string' && idVal.trim())) {
                                                        const s = String(idVal);
                                                        displayValue = (userCache && userCache.has(s)) ? userCache.get(s)! : s;
                                                    } else if (typeof rawVal === 'string' && rawVal.trim()) {
                                                        displayValue = rawVal;
                                                    }
                                                }
                                            } else {
                                                const val = fieldName ? item[fieldName] : undefined;
                                                displayValue = (val === null || typeof val === 'undefined') ? '' : String(val);
                                            }

                                            // Default cell render: value only (actions are rendered in the actions column)
                                            return (
                                                <div>{displayValue}</div>
                                            );
                                        }}
                                    />
                                )}
                            </div>
                        </PivotItem>
                        <PivotItem headerText="Config" itemKey="config">
                            <div style={{ marginTop: 12 }}>
                                {/* Schema Viewer moved into Config tab */}
                                {fields && (
                                    <div style={{ marginTop: 18 }}>
                                        <h4>List Fields ({fields.length})</h4>
                                        <div style={{ maxHeight: 220, overflow: 'auto', border: '1px solid #e1e1e1', padding: 8 }}>
                                            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                                <thead>
                                                    <tr>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>InternalName</th>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>Title</th>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>Type</th>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>Hidden</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {fields.map((f) => (
                                                        <tr key={f.InternalName}>
                                                            <td style={{ padding: 6 }}>{f.InternalName}</td>
                                                            <td style={{ padding: 6 }}>{f.Title}</td>
                                                            <td style={{ padding: 6 }}>{f.TypeAsString}</td>
                                                            <td style={{ padding: 6 }}>{f.Hidden ? 'Yes' : 'No'}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}

                                {/* Views viewer moved into Config tab */}
                                {views && (
                                    <div style={{ marginTop: 18 }}>
                                        <h4>List Views ({views.length})</h4>
                                        <div style={{ marginBottom: 8 }}>
                                            <DefaultButton
                                                text={showRawJson ? 'Hide Raw JSON' : 'Show Raw JSON'}
                                                onClick={() => setShowRawJson(!showRawJson)}
                                            />
                                        </div>
                                        <div style={{ maxHeight: 220, overflow: 'auto', border: '1px solid #e1e1e1', padding: 8 }}>
                                            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                                <thead>
                                                    <tr>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>Title</th>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>Id</th>
                                                        <th style={{ textAlign: 'left', padding: 6 }}>ViewFields</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {views.map((v) => (
                                                        <tr key={v.Id}>
                                                            <td style={{ padding: 6 }}>{v.Title}</td>
                                                            <td style={{ padding: 6 }}>{v.Id}</td>
                                                            <td style={{ padding: 6 }}>{Array.isArray(v.ViewFields) ? v.ViewFields.join(', ') : String(v.ViewFields)}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                        {showRawJson && (
                                            <div style={{ marginTop: 12 }}>
                                                <div style={{ whiteSpace: 'pre', background: '#f3f2f1', padding: 10, borderRadius: 4, maxHeight: 360, overflow: 'auto' }}>
                                                    <pre style={{ margin: 0 }}>{JSON.stringify({ views, fields }, null, 2)}</pre>
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                )}
                            </div>
                        </PivotItem>
                    </Pivot>
                </div>
            </div>

            {/* Schema and Views are shown in the Config tab only (duplicates removed) */}

            <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)} isBlocking={false}>
                <div style={{ padding: 18, maxWidth: 1000, maxHeight: '80vh', overflow: 'auto' }}>
                    <h3>{isEditing ? 'Edit Property' : 'New Property'}</h3>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 12, alignItems: 'start' }}>
                        {getViewFields(selectedView).map((field: string) => {
                            const meta = getMetaForField(field);
                            const rawType = meta && (meta.TypeAsString || meta.FieldType || meta.Type) ? String(meta.TypeAsString || meta.FieldType || meta.Type) : '';
                            const type = rawType.toLowerCase();
                            const spanTwo = type.includes('note') || type.includes('multiline');
                            return (
                                <div key={field} style={spanTwo ? { gridColumn: 'span 2' } : undefined}>
                                    {renderFieldControl(field)}
                                </div>
                            );
                        })}

                        <div style={{ gridColumn: 'span 2', display: 'flex', gap: 8, marginTop: 6 }}>
                            <PrimaryButton text="Save" onClick={save} />
                            <DefaultButton text="Cancel" onClick={() => setIsModalOpen(false)} />
                        </div>
                    </div>
                </div>
            </Modal>
            <Modal isOpen={isBulkModalOpen} onDismiss={() => setIsBulkModalOpen(false)} isBlocking={false}>
                <div style={{ padding: 18, maxWidth: 900, maxHeight: '80vh', overflow: 'auto' }}>
                    <h3>Bulk Update Fields</h3>
                    <div style={{ marginBottom: 12 }}>
                        <div style={{ marginBottom: 8 }}>Select fields to update across all records</div>
                        <Dropdown
                            placeholder="Select fields"
                            multiSelect
                            options={(fields || []).map((f: any) => ({ key: String(f.InternalName), text: f.Title || f.InternalName }))}
                            selectedKeys={bulkSelectedFields}
                            onChange={(_, option) => {
                                if (!option) return;
                                const key = String(option.key);
                                const cur = Array.isArray(bulkSelectedFields) ? [...bulkSelectedFields] : [];
                                if (option.selected) {
                                    if (cur.indexOf(key) === -1) cur.push(key);
                                } else {
                                    const idx = cur.indexOf(key);
                                    if (idx >= 0) cur.splice(idx, 1);
                                }
                                setBulkSelectedFields(cur);
                            }}
                            styles={{ root: { minWidth: '100%' } }}
                        />
                    </div>

                    {bulkSelectedFields && bulkSelectedFields.length > 0 && (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 12 }}>
                            {bulkSelectedFields.map((f) => (
                                <div key={f}>{renderBulkFieldControl(f)}</div>
                            ))}
                        </div>
                    )}

                    {bulkProgress && (
                        <div style={{ marginTop: 12 }}>
                            <div>Progress: {bulkProgress.done} / {bulkProgress.total}</div>
                            <div style={{ marginTop: 6 }}>{bulkLoading && <Spinner size={SpinnerSize.small} />}</div>
                        </div>
                    )}

                    <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                        <PrimaryButton
                            text="Apply to all"
                            onClick={() => {
                                const count = items ? items.length : 0;
                                if (!bulkSelectedFields || bulkSelectedFields.length === 0) {
                                    setError('No fields selected for bulk update');
                                    return;
                                }
                                if (!confirm(`Apply changes to all ${count} items? This cannot be undone.`)) return;
                                void performBulkUpdate();
                            }}
                            disabled={bulkLoading || !bulkSelectedFields || bulkSelectedFields.length === 0}
                        />
                        <DefaultButton text="Cancel" onClick={() => setIsBulkModalOpen(false)} />
                    </div>
                </div>
            </Modal>
        </div>
    );
};

export default PropertyManager;
