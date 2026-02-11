import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './Leaderboard.module.scss';
import { IconButton, Spinner, SpinnerSize, Pivot, PivotItem } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyService, IPropertyItem } from '../../propertyManager/services/PropertyService';

export interface ILeaderboardProps {
    context: WebPartContext;
    title?: string;
    subtitle?: string;
    listName?: string; // name of SharePoint list (default KPI)
    categoryField?: string; // field used to segment categories (default 'Category')
    slideTitles?: {
        [key: string]: {
            title?: string;
            subtitle?: string;
        };
    };
    categoryHeading?: string;
    categorySubheading?: string;
    categorySlideTitle?: string;
    categorySlideSubtitle?: string;
}

const DEFAULTS = {
    listName: 'KPI',
    categoryField: 'Category'
};

export const Leaderboard: React.FC<ILeaderboardProps> = (props) => {
    const { context, title, subtitle, listName = DEFAULTS.listName, categoryField = DEFAULTS.categoryField } = props;
    const { slideTitles = {}, categoryHeading = '', categorySubheading = '', categorySlideTitle = '', categorySlideSubtitle = '' } = props;
    const [items, setItems] = useState<IPropertyItem[] | null>(null);
    const [loading, setLoading] = useState<boolean>(false);
    const [currentIndex, setCurrentIndex] = useState<number>(0);
    const [categories, setCategories] = useState<string[]>([]);

    useEffect(() => {
        if (!context) return;
        PropertyService.init(context);
        const load = async () => {
            setLoading(true);
            try {
                // request Employee and explicitly include Employee/JobTitle so JobTitle is returned when available
                const data = await PropertyService.getItemsFromList(listName, undefined, 500, ['Employee', 'Employee/JobTitle']);
                setItems(Array.isArray(data) ? data : []);
                // debug: log sample item to help diagnose missing person/photo fields
                // eslint-disable-next-line no-console
                console.debug('Leaderboard: sample item (for debugging)', Array.isArray(data) && data.length ? data[0] : data);

                // additional diagnostics: gather job-title candidate fields from each item
                try {
                    const jobCandidates = ['JobTitle', 'jobTitle', 'Position', 'Designation', 'OData__JobTitle', 'SPSJobTitle', 'SPS-JobTitle', 'Job Title'];
                    const diagnostics = (Array.isArray(data) ? data : []).map((d: any) => {
                        const found: any = {};
                        for (const k of jobCandidates) {
                            if (d && Object.prototype.hasOwnProperty.call(d, k) && d[k]) found[k] = d[k];
                        }
                        for (const key of Object.keys(d || {})) {
                            const val = d[key];
                            if (val && typeof val === 'object') {
                                for (const k of jobCandidates) {
                                    if (Object.prototype.hasOwnProperty.call(val, k) && val[k]) found[`${key}.${k}`] = val[k];
                                }
                                if (val.Value && typeof val.Value === 'object') {
                                    for (const k of jobCandidates) {
                                        if (Object.prototype.hasOwnProperty.call(val.Value, k) && val.Value[k]) found[`${key}.Value.${k}`] = val.Value[k];
                                    }
                                }
                            }
                        }
                        return { id: d && (d.Id || d.ID || d.id) || null, title: d && (d.Title || d.Name || ''), candidates: found };
                    });
                    // eslint-disable-next-line no-console
                    console.debug('Leaderboard: job-title diagnostics', diagnostics);
                } catch (err) {
                    // eslint-disable-next-line no-console
                    console.warn('Leaderboard: diagnostics failed', err);
                }
                // derive categories from the configured category field
                const cats = new Set<string>();
                (data || []).forEach((d: any) => {
                    const v = d[categoryField];
                    if (v) cats.add(String(v));
                });
                const catArr = Array.from(cats);
                if (catArr.length === 0) catArr.push('All');
                setCategories(catArr);
                setCurrentIndex(0);
            } catch (e) {
                // eslint-disable-next-line no-console
                console.error('Leaderboard load error', e);
                setItems([]);
            } finally {
                setLoading(false);
            }
        };
        void load();
    }, [context, listName, categoryField]);

    const goPrev = () => setCurrentIndex(i => Math.max(0, i - 1));
    const goNext = () => setCurrentIndex(i => Math.min((categories || []).length - 1, i + 1));

    const curCategory = categories && categories.length > 0 ? categories[currentIndex] : undefined;

    // Helper to sanitize category key (must match web part _sanitizeKey)
    const sanitizeKey = (value?: string) => String(value || '').replace(/[^a-zA-Z0-9]/g, '_');

    const curSlideTitle = curCategory ? slideTitles[sanitizeKey(curCategory)] : undefined;
    // prefer slide-specific title, otherwise generate a category-aware title (e.g. "Leaderboard-Individual"),
    // falling back to the configured global title
    const headerTitleText = (curSlideTitle && curSlideTitle.title)
        || (curCategory && curCategory !== 'All' ? `${title || 'Leaderboard'}-${curCategory}` : title || 'Leaderboard');

    // prefer slide-specific subtitle, otherwise generate a category-aware subtitle when on a category slide,
    // else fall back to the configured global subtitle or the categorySlideSubtitle if provided
    const headerSubtitleText = (curSlideTitle && curSlideTitle.subtitle)
        || (curCategory && curCategory !== 'All'
            ? ((subtitle && String(subtitle).trim().length > 0) ? `${subtitle}-${curCategory}` : (categorySlideSubtitle || ''))
            : (subtitle || categorySlideSubtitle || ''));

    const filtered = (items || []).filter((it: any) => {
        if (!curCategory || curCategory === 'All') return true;
        const v = it[categoryField];
        return String(v) === String(curCategory);
    });

    // Sort by Rank (ascending) if present, otherwise by Score/Value (descending)
    const sorted = filtered.slice().sort((a: any, b: any) => {
        const aRank = Number(a.Rank ?? a.Ranking ?? a.RankValue ?? NaN);
        const bRank = Number(b.Rank ?? b.Ranking ?? b.RankValue ?? NaN);
        if (!isNaN(aRank) && !isNaN(bRank)) {
            return aRank - bRank; // ascending rank: 1,2,3...
        }
        const aScore = Number(a.Score ?? a.Value ?? 0);
        const bScore = Number(b.Score ?? b.Value ?? 0);
        return bScore - aScore; // fallback: highest score first
    });

    // Use the full sorted list (display all rows for the selected category)
    const topList = sorted;

    // Attempt to extract an email address from an item using common field names
    const getEmailFromItem = (it: any): string | null => {
        if (!it) return null;
        const tryKeys = ['EMail'];
        for (const k of tryKeys) {
            const v = it[k];
            if (!v) continue;
            if (typeof v === 'string' && v.indexOf('@') > -1) return v;
            if (typeof v === 'object') {
                if (v.Email && String(v.Email).indexOf('@') > -1) return v.Email;
                if (v.EMail && String(v.EMail).indexOf('@') > -1) return v.EMail;
                if (v.LoginName && String(v.LoginName).indexOf('@') > -1) {
                    const parts = String(v.LoginName).split('|');
                    return parts.pop() ?? null;
                }
            }
        }

        // Look for person-like fields where the value may be an object with Email
        for (const key of Object.keys(it)) {
            const val = it[key];
            if (val && typeof val === 'object') {
                if (val.Email && String(val.Email).indexOf('@') > -1) return val.Email;
                if (val.EMail && String(val.EMail).indexOf('@') > -1) return val.EMail;
            }
        }

        return null;
    };

    // Attempt to extract a Job Title from common fields on the item or person object
    const getJobTitle = (it: any): string | null => {
        if (!it) return null;

        // helper: search for explicit job-title keys and avoid generic `Title` fallback
        const jobKeyRx = /(jobtitle|job|position|designation|role)/i; // do NOT match plain 'title'
        const maxDepth = 3;

        const personCandidates = ['JobTitle', 'jobTitle', 'OData__JobTitle', 'SPSJobTitle', 'SPS-JobTitle'];

        const findJob = (obj: any, depth = 0): string | null => {
            if (!obj || typeof obj !== 'object' || depth > maxDepth) return null;
            for (const key of Object.keys(obj)) {
                const val = obj[key];
                if (val === undefined || val === null) continue;

                // Prioritise clearly job-related field names
                if (jobKeyRx.test(key)) {
                    if (typeof val === 'string' && val.trim() !== '') return val.trim();
                    if (typeof val === 'object') {
                        if (typeof val.Value === 'string' && val.Value.trim() !== '') return val.Value.trim();
                        // avoid returning person 'Title' (display name) here
                        if (typeof val.JobTitle === 'string' && val.JobTitle.trim() !== '') return val.JobTitle.trim();
                    }
                }

                // If this property looks like a person object, check common explicit job properties
                if (typeof val === 'object') {
                    for (const pc of personCandidates) {
                        if (val[pc] && typeof val[pc] === 'string' && val[pc].trim() !== '') return val[pc].trim();
                    }
                    if (val.Value && typeof val.Value === 'object') {
                        for (const pc of personCandidates) {
                            if (val.Value[pc] && typeof val.Value[pc] === 'string' && val.Value[pc].trim() !== '') return val.Value[pc].trim();
                        }
                    }
                }

                // recurse into nested objects
                if (typeof val === 'object') {
                    const found = findJob(val, depth + 1);
                    if (found) return found;
                }
            }
            return null;
        };

        // 1) try direct, explicit job-like keys on the item
        const explicit = findJob(it, 0);
        if (explicit) return String(explicit).trim();

        // 2) check common person fields specifically (Employee, AssignedTo, Author, Editor)
        const personKeys = ['Employee', 'employee', 'AssignedTo', 'Author', 'Editor', 'Assigned_x0020_To'];
        for (const pk of personKeys) {
            const p = it[pk];
            if (!p) continue;
            const persons = Array.isArray(p) ? p : [p];
            for (const person of persons) {
                if (!person || typeof person !== 'object') continue;
                for (const pc of personCandidates) {
                    if (person[pc] && typeof person[pc] === 'string' && person[pc].trim() !== '') return person[pc].trim();
                }
                // nested pattern
                if (person.Value && typeof person.Value === 'object') {
                    for (const pc of personCandidates) {
                        if (person.Value[pc] && typeof person.Value[pc] === 'string' && person.Value[pc].trim() !== '') return person.Value[pc].trim();
                    }
                }
            }
        }

        // 3) last-resort: avoid returning a generic 'Title' (display name). Return null instead.
        return null;
    };

    // Build an avatar/photo URL for the item.
    // Prefer explicit photo fields, then look for an `Employee` person field (or other person-like fields)
    // and use the user's email to construct the SharePoint user photo URL.
    const getAvatarUrl = (it: any): string | null => {
        if (!it) return null;

        // Quick checks for explicit photo fields present on the item
        const photoCandidates = ['PhotoUrl', 'Picture', 'Image', 'Photo', 'UserPhoto', 'PictureUrl'];
        for (const k of photoCandidates) {
            const v = it[k];
            if (!v) continue;
            if (typeof v === 'string' && v.trim() !== '') return v;
            if (typeof v === 'object') {
                if (v.Url) return v.Url;
                if (v.url) return v.url;
            }
        }

        // Prefer a specifically-named Employee field if present (user-type field)
        const empKeys = ['Employee', 'employee', 'AssignedTo', 'Author', 'Editor', 'Assigned_x0020_To'];
        for (const k of empKeys) {
            const candidate = it[k];
            if (!candidate) continue;
            // candidate might be an object or array (people picker can be multi-valued)
            const persons = Array.isArray(candidate) ? candidate : [candidate];
            for (const p of persons) {
                if (!p) continue;
                // common shapes: { Email, EMail, LoginName, Title }
                const email = (p.Email || p.EMail || p.Mail) || (p.LoginName && String(p.LoginName).indexOf('@') > -1 ? String(p.LoginName).split('|').pop() : null);
                if (email && String(email).indexOf('@') > -1) {
                    try {
                        const base = context?.pageContext?.web?.absoluteUrl || '';
                        // use SharePoint userphoto handler
                        const url = `${base}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(String(email))}`;
                        return url;
                    } catch (e) {
                        // fallback to empty
                    }
                }
            }
        }

        // If Employee wasn't found, scan object fields for a person-like object with Email
        for (const key of Object.keys(it)) {
            const val = it[key];
            if (!val || typeof val !== 'object') continue;
            const email = (val.Email || val.EMail || val.Mail) || (val.LoginName && String(val.LoginName).indexOf('@') > -1 ? String(val.LoginName).split('|').pop() : null);
            if (email && String(email).indexOf('@') > -1) {
                const base = context?.pageContext?.web?.absoluteUrl || '';
                return `${base}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(String(email))}`;
            }
            // check nested structures
            if (val && typeof val === 'object') {
                if (val.Value && typeof val.Value === 'object') {
                    const nested = val.Value;
                    const email2 = (nested.Email || nested.EMail || nested.Mail) || (nested.LoginName && String(nested.LoginName).indexOf('@') > -1 ? String(nested.LoginName).split('|').pop() : null);
                    if (email2 && String(email2).indexOf('@') > -1) {
                        const base = context?.pageContext?.web?.absoluteUrl || '';
                        return `${base}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(String(email2))}`;
                    }
                }
            }
        }

        return null;
    };

    const handleContact = (it: any) => {
        const email = getEmailFromItem(it);
        const displayName = it && (it.Title || it.Name || it.Title0 || it.DisplayName) ? (it.Title || it.Name || it.Title0 || it.DisplayName) : '';
        if (email) {
            // Prefer opening a Teams chat deep link; fall back to mailto if Teams is blocked
            const teamsUrl = `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`;
            try {
                window.open(teamsUrl, '_blank');
                return;
            } catch (e) {
                // ignore and fallback
            }
            // fallback to mailto
            const subject = encodeURIComponent(`Message for ${displayName || 'a colleague'}`);
            const mailto = `mailto:${encodeURIComponent(email)}?subject=${subject}`;
            window.location.href = mailto;
            return;
        }

        // No email found — open a mailto with no recipient but prefilled subject mentioning the person
        const subj = encodeURIComponent(`Message regarding ${displayName || 'this person'}`);
        window.location.href = `mailto:?subject=${subj}`;
    };

    return (
        <div className={styles.leaderboard}>
            <div className={styles.header}>
                <div style={{ width: 48, height: 48, display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: 8, background: 'rgba(255,255,255,0.12)' }}>
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M7 3V5" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /><path d="M17 3V5" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /><path d="M7 21a4 4 0 0 1-4-4v-4h6v8zM21 21a4 4 0 0 0 4-4v-4h-6v8z" stroke="#fff" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" /></svg>
                </div>
                <div style={{ flex: 1 }}>
                    <div className={styles.headerTitle}>{headerTitleText}</div>
                    {(headerSubtitleText || categoryHeading) && <div className={styles.headerSubtitle}>{headerSubtitleText || categoryHeading}</div>}
                    {categorySubheading && <div className={styles.headerSubtitle} style={{ fontSize: 12, opacity: 0.9 }}>{categorySubheading}</div>}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 8 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Prev" onClick={goPrev} disabled={currentIndex === 0} />
                        <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next" onClick={goNext} disabled={currentIndex >= (categories.length - 1)} />
                    </div>
                    <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.95)' }}>{curCategory || ''}</div>
                </div>
            </div>

            <div className={styles.list}>
                {loading && <Spinner size={SpinnerSize.large} label="Loading..." />}
                {!loading && topList && topList.length === 0 && (
                    <div style={{ padding: 16 }}>No data</div>
                )}
                {!loading && topList.map((it, idx) => {
                    const titleText = it.Title || it.Title0 || it.Name || `Item ${it.Id || idx + 1}`;
                    const photo = getAvatarUrl(it) || it.PhotoUrl || it.Picture || it.Image || (it.Photo && it.Photo.Url) || '';
                    const rankVal = (it.Rank ?? it.Ranking ?? it.RankValue);
                    const rankDisplay = rankVal !== undefined && rankVal !== null && rankVal !== '' ? String(rankVal) : String(idx + 1);

                    // Robust score extraction: try many common field names, handle strings with '%', and fractions (0-1)
                    const getNumericScore = (row: any): number | null => {
                        if (!row) return null;
                        const keys = ['Score', 'Scores', 'Value', 'Values', 'Percentage', 'Percent', 'ScoreValue', 'Result', 'Score_x0020_', 'Scores_x0020_'];
                        for (const k of keys) {
                            const v = row[k];
                            if (v === undefined || v === null) continue;
                            if (typeof v === 'number') return v;
                            if (typeof v === 'string') {
                                const s = v.trim();
                                // strip percent sign
                                if (s.endsWith('%')) {
                                    const n = parseFloat(s.replace('%', '').replace(/,/g, ''));
                                    if (!isNaN(n)) return n;
                                }
                                // numeric string
                                const n = parseFloat(s.replace(/,/g, ''));
                                if (!isNaN(n)) return n;
                            }
                            if (typeof v === 'object') {
                                // common nested structures
                                if (v.Value !== undefined && v.Value !== null) {
                                    const n = Number(v.Value);
                                    if (!isNaN(n)) return n;
                                }
                                if (v.Score !== undefined && v.Score !== null) {
                                    const n = Number(v.Score);
                                    if (!isNaN(n)) return n;
                                }
                                if (v.Percentage !== undefined && v.Percentage !== null) {
                                    const n = Number(v.Percentage);
                                    if (!isNaN(n)) return n;
                                }
                            }
                        }

                        // last-resort: scan all fields for a percent string
                        for (const k of Object.keys(row)) {
                            const v = row[k];
                            if (typeof v === 'string' && v.indexOf('%') !== -1) {
                                const n = parseFloat(v.replace('%', '').replace(/,/g, ''));
                                if (!isNaN(n)) return n;
                            }
                        }

                        return null;
                    };

                    const rawScore = getNumericScore(it);
                    let displayScore = '-';
                    if (rawScore !== null) {
                        let normalized = rawScore;
                        // if fraction like 0.83, convert to percent
                        if (normalized > 0 && normalized <= 1) normalized = normalized * 100;
                        // clamp and format
                        if (!isNaN(normalized)) {
                            displayScore = `${Number(normalized.toFixed(2))}%`;
                        }
                    }

                    return (
                        <div key={it.Id || idx} className={styles.item}>
                            <div className={styles.itemLeft} style={{ alignItems: 'center' }}>
                                <div style={{ width: 36, textAlign: 'center', color: '#ffffffcc', fontWeight: 700 }}>{rankDisplay}</div>
                                <div className={styles.avatar} role="button" title={`Message ${titleText}`} onClick={() => handleContact(it)} style={{ cursor: 'pointer' }}>
                                    {photo ? <img src={photo} alt={titleText} /> : <div style={{ color: '#0a72c6', fontWeight: 700 }}>{String(titleText || '').charAt(0)}</div>}
                                </div>
                                <div className={styles.itemText} style={{ marginLeft: 8 }}>
                                    <div className={styles.itemTitle} onClick={() => handleContact(it)} style={{ cursor: 'pointer' }}>{titleText}</div>
                                    {(() => {
                                        const jt = getJobTitle(it);
                                        return jt ? <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.75)' }}>{jt}</div> : null;
                                    })()}
                                </div>
                            </div>

                            <div style={{ minWidth: 100, textAlign: 'right' }}>
                                <div className={styles.score}>{displayScore}</div>
                            </div>
                        </div>
                    );
                })}
            </div>

            <div className={styles.controls}>
                <div className={styles.carouselDots}>
                    {(categories || []).map((c, i) => <div key={c + i} className={`${styles.dot} ${i === currentIndex ? styles.active : ''}`} onClick={() => setCurrentIndex(i)} />)}
                </div>
            </div>
        </div>
    );
};

export default Leaderboard;
