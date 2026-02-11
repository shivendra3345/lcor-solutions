import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
    type IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneButton
} from '@microsoft/sp-property-pane';

import Leaderboard from './components/Leaderboard';
import { PropertyService } from '../propertyManager/services/PropertyService';

export interface ILeaderboardWebPartProps {
    title?: string;
    subtitle?: string;
    listName?: string;
    categoryField?: string;
    slideTitles?: {
        [key: string]: {
            title?: string;
            subtitle?: string;
        };
    };
    categorySlideTitle?: string;
    categorySlideSubtitle?: string;
    // per-category (choice) titles — configured from the property pane
    individualSlideTitle?: string;
    individualSlideSubtitle?: string;
    propertySlideTitle?: string;
    propertySlideSubtitle?: string;
}

export default class LeaderboardWebPart extends BaseClientSideWebPart<ILeaderboardWebPartProps> {
    private _categoryFields: any[] = [];
    protected onInit(): Promise<void> {
        PropertyService.init(this.context);
        return super.onInit();
    }

    public render(): void {
        const element: React.ReactElement = React.createElement(Leaderboard, {
            context: this.context,
            title: this.properties.title || 'Leaderboard',
            subtitle: this.properties.subtitle || '',
            listName: this.properties.listName || 'KPI',
            categoryField: this.properties.categoryField || 'Category'
            ,
            // build a slideTitles map that includes any configured per-category titles
            slideTitles: Object.assign({}, this.properties.slideTitles || {}, {
                [this._sanitizeKey('Individual')]: {
                    title: (this.properties as any).individualSlideTitle || undefined,
                    subtitle: (this.properties as any).individualSlideSubtitle || undefined
                },
                [this._sanitizeKey('Property')]: {
                    title: (this.properties as any).propertySlideTitle || undefined,
                    subtitle: (this.properties as any).propertySlideSubtitle || undefined
                }
            }),
            categoryHeading: (this.properties as any).categoryHeading || '',
            categorySubheading: (this.properties as any).categorySubheading || '',
            categorySlideTitle: (this.properties as any).categorySlideTitle || '',
            categorySlideSubtitle: (this.properties as any).categorySlideSubtitle || ''
        } as any);

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected async onPropertyPaneConfigurationStart(): Promise<void> {
        // no dynamic slide title fields anymore; keep property pane minimal
        this._categoryFields = [];
    }

    private async buildCategoryFields(): Promise<void> {
        this._categoryFields = [];

        const listName = this.properties.listName || 'KPI';
        const categoryField = this.properties.categoryField || 'Category';

        try {
            const items = await PropertyService.getItemsFromList(listName, undefined, 500, ['Employee']);
            // Build category-based slide title fields (prefix keys with 'category_')
            const cats = Array.from(new Set((items || []).map((it: any) => String(it[categoryField] || '').trim()).filter((c: string) => c)));
            cats.forEach((cat: string) => {
                const key = `category_${this._sanitizeKey(cat)}`;

                this._categoryFields.push(
                    PropertyPaneTextField(`slideTitles.${key}.title`, {
                        label: `Category: ${cat} — Slide Title`,
                        placeholder: `Title for category ${cat}`
                    })
                );

                this._categoryFields.push(
                    PropertyPaneTextField(`slideTitles.${key}.subtitle`, {
                        label: `Category: ${cat} — Slide Subtitle`,
                        placeholder: `Subtitle for category ${cat}`
                    })
                );
            });

            // Build individual-based slide title fields (prefix keys with 'individual_')
            // Use the expanded Employee person field when available to list unique people.
            const persons = new Map<string, string>(); // key -> displayName
            (items || []).forEach((it: any) => {
                const emp = it.Employee || it.employee || it.AssignedTo || it.Author || it.Editor;
                if (!emp) return;
                const personsArr = Array.isArray(emp) ? emp : [emp];
                personsArr.forEach((p: any) => {
                    if (!p) return;
                    const name = p.Title || p.Title0 || p.Name || (p.Email ? p.Email : '') || (p.EMail ? p.EMail : '');
                    const display = String(name || '').trim();
                    if (!display) return;
                    const key = this._sanitizeKey(display);
                    if (!persons.has(key)) persons.set(key, display);
                });
            });

            persons.forEach((display, k) => {
                const key = `individual_${k}`;
                this._categoryFields.push(
                    PropertyPaneTextField(`slideTitles.${key}.title`, {
                        label: `Individual: ${display} — Slide Title`,
                        placeholder: `Title for ${display}`
                    })
                );
                this._categoryFields.push(
                    PropertyPaneTextField(`slideTitles.${key}.subtitle`, {
                        label: `Individual: ${display} — Slide Subtitle`,
                        placeholder: `Subtitle for ${display}`
                    })
                );
            });

            // if no categories found, add a refresh button so user can retry
            if (this._categoryFields.length === 0) {
                this._categoryFields.push(PropertyPaneButton('refreshCategories', {
                    text: 'No categories found — Refresh',
                    buttonType: 0,
                    onClick: async () => {
                        await this.buildCategoryFields();
                        this.context.propertyPane.refresh();
                    }
                }));
            }
        } catch (e) {
            // on error, provide a refresh button
            this._categoryFields = [PropertyPaneButton('refreshCategories', {
                text: 'Refresh categories',
                buttonType: 0,
                onClick: async () => {
                    await this.buildCategoryFields();
                    this.context.propertyPane.refresh();
                }
            })];
        }
    }

    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
        // support nested slideTitles.<key>.title and .subtitle
        if (propertyPath && propertyPath.indexOf('slideTitles.') === 0) {
            const rest = propertyPath.replace('slideTitles.', '');
            const parts = rest.split('.'); // [key, title|subtitle]
            if (parts.length >= 2) {
                const key = parts[0];
                const field = parts[1];
                this.properties.slideTitles = (this.properties as any).slideTitles || {};
                (this.properties as any).slideTitles[key] = (this.properties as any).slideTitles[key] || {};
                (this.properties as any).slideTitles[key][field] = newValue || '';
                this.render();
                return;
            }
        }

        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }

    private _sanitizeKey(value: string): string {
        return String(value || '').replace(/[^a-zA-Z0-9]/g, '_');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: 'Leaderboard settings'
                    },
                    groups: [
                        {
                            groupName: 'Settings',
                            groupFields: [
                                PropertyPaneTextField('listName', { label: 'List name', placeholder: 'e.g., KPI' }),
                                PropertyPaneTextField('individualSlideTitle', { label: 'Individual slide title', placeholder: 'Title when Category = Individual' }),
                                PropertyPaneTextField('individualSlideSubtitle', { label: 'Individual slide subtitle', placeholder: 'Subtitle when Category = Individual' }),
                                PropertyPaneTextField('propertySlideTitle', { label: 'Property slide title', placeholder: 'Title when Category = Property' }),
                                PropertyPaneTextField('propertySlideSubtitle', { label: 'Property slide subtitle', placeholder: 'Subtitle when Category = Property' })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
