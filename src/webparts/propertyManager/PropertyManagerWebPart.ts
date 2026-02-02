import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PropertyManager from './components/PropertyManager';
import { PropertyService } from './services/PropertyService';

export interface IPropertyManagerWebPartProps { }

export default class PropertyManagerWebPart extends BaseClientSideWebPart<IPropertyManagerWebPartProps> {
    protected onInit(): Promise<void> {
        // Initialize the PropertyService with SPFx context
        PropertyService.init(this.context);
        return super.onInit();
    }

    public render(): void {
        const element: React.ReactElement = React.createElement(PropertyManager, {
            context: this.context
        } as any);

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }
}
