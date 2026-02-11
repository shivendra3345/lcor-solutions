import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';

import QuickLinks from './components/QuickLinks';
import { PropertyService } from '../propertyManager/services/PropertyService';

export interface IQuickLinksWebPartProps {
    listName?: string;
    pageTitle?: string;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
    protected onInit(): Promise<void> {
        PropertyService.init(this.context);
        return super.onInit();
    }

    public render(): void {
        const element: React.ReactElement = React.createElement(QuickLinks, {
            context: this.context,
            listName: this.properties.listName || 'QuickLinks',
            pageTitle: this.properties.pageTitle || 'Quick Links'
        } as any);

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: { description: 'Quick Links settings' },
                    groups: [
                        {
                            groupName: 'Settings',
                            groupFields: [
                                PropertyPaneTextField('listName', { label: 'List name', placeholder: 'e.g., QuickLinks' }),
                                PropertyPaneTextField('pageTitle', { label: 'Page title', placeholder: 'Quick Links' })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
