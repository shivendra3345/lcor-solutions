import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './QuickLinks.module.scss';
import { IconButton, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyService } from '../../propertyManager/services/PropertyService';

export interface IQuickLinksProps {
    context: WebPartContext;
    listName?: string;
    pageTitle?: string;
    pageSize?: number;
}

const PAGE_SIZE = 10;

const QuickLinks: React.FC<IQuickLinksProps> = (props) => {
    const { context, listName = 'QuickLinks', pageTitle = 'Quick Links', pageSize: pageSizeProp } = props;
    const [items, setItems] = useState<any[] | null>(null);
    const [loading, setLoading] = useState<boolean>(false);
    const [pageIndex, setPageIndex] = useState<number>(0);

    useEffect(() => {
        if (!context) return;
        PropertyService.init(context);
        const load = async () => {
            setLoading(true);
            try {
                const data = await PropertyService.getItemsFromList(listName, undefined, 500, []);
                setItems(Array.isArray(data) ? data : []);
                setPageIndex(0);
            } catch (e) {
                // eslint-disable-next-line no-console
                console.error('QuickLinks load error', e);
                setItems([]);
            } finally {
                setLoading(false);
            }
        };
        void load();
    }, [context, listName]);

    const pageSize = (typeof pageSizeProp === 'number' && pageSizeProp > 0) ? Math.floor(pageSizeProp) : PAGE_SIZE;
    const total = (items || []).length;
    const pages = Math.max(1, Math.ceil(total / pageSize));

    const goPrev = () => setPageIndex(i => Math.max(0, i - 1));
    const goNext = () => setPageIndex(i => Math.min(pages - 1, i + 1));

    const pageItems = (items || []).slice(pageIndex * pageSize, (pageIndex + 1) * pageSize);

    // Ensure pageIndex is within bounds when items or pageSize change
    React.useEffect(() => {
        if (pageIndex > pages - 1) {
            setPageIndex(Math.max(0, pages - 1));
        }
    }, [pages, pageIndex]);

    const extractLink = (it: any): { url?: string; desc?: string } => {
        // Common SharePoint URL field shapes: { Url, Description } or simple string
        const candidates = ['URL', 'Url', 'Link', 'Hyperlink', 'HyperlinkUrl', 'Address'];
        for (const k of candidates) {
            const v = it[k];
            if (!v) continue;
            if (typeof v === 'string') return { url: v, desc: it.Title || '' };
            if (typeof v === 'object') {
                if (v.Url || v.url) return { url: v.Url || v.url, desc: v.Description || v.Description || it.Title || '' };
            }
        }
        // fallback: common pair fields
        if (it.Link && typeof it.Link === 'string') return { url: it.Link, desc: it.Title || '' };
        if (it.URL0 && typeof it.URL0 === 'string') return { url: it.URL0, desc: it.Title || '' };
        return { url: undefined, desc: it.Title || '' };
    };

    return (
        <div className={styles.quicklinks}>
            <div className={styles.header}>
                <div style={{ width: 48, height: 48, display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: 8, background: 'rgba(255,255,255,0.12)' }}>
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 14l-4-4 4-4" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /><path d="M14 10l4 4-4 4" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
                </div>
                <div style={{ flex: 1 }}>
                    <div className={styles.headerTitle}>{pageTitle}</div>
                    <div className={styles.headerSubtitle}>{total} links</div>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 8 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Prev" onClick={goPrev} disabled={pageIndex === 0} />
                        <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next" onClick={goNext} disabled={pageIndex >= (pages - 1)} />
                    </div>
                    <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.95)' }}>{`Page ${pageIndex + 1} / ${pages}`}</div>
                </div>
            </div>

            <div className={styles.list}>
                {loading && <Spinner size={SpinnerSize.large} label="Loading..." />}
                {!loading && pageItems && pageItems.length === 0 && (
                    <div style={{ padding: 16 }}>No links</div>
                )}

                {!loading && pageItems.map((it, idx) => {
                    const link = extractLink(it);
                    const title = it.Title || link.desc || `Link ${(pageIndex * PAGE_SIZE) + idx + 1}`;
                    const url = link.url || '#';
                    return (
                        <div key={(it.Id || it.ID) ?? `${pageIndex}-${idx}`} className={styles.item}>
                            <div className={styles.itemLeft}>
                                <div className={styles.itemTitle}><a href={url} target="_blank" rel="noreferrer">{title}</a></div>
                                {it.Description && <div>{it.Description}</div>}
                            </div>
                            <div style={{ minWidth: 60, textAlign: 'right' }}>
                                <IconButton iconProps={{ iconName: 'OpenInNewWindow' }} title="Open" onClick={() => window.open(url, '_blank')} />
                            </div>
                        </div>
                    );
                })}
            </div>

            <div className={styles.controls}>
                <div className={styles.carouselDots}>
                    {Array.from({ length: pages }).map((_, i) => <div key={i} className={`${styles.dot} ${i === pageIndex ? styles.active : ''}`} onClick={() => setPageIndex(i)} />)}
                </div>
            </div>
        </div>
    );
};

export default QuickLinks;
