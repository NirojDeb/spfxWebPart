import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxWebPartDemoNjWebPartStrings';
import SpfxWebPartDemoNj from './components/SpfxWebPartDemoNj';
import { ISpfxWebPartDemoNjProps } from './components/ISpfxWebPartDemoNjProps';
import {sp} from '@pnp/sp';

export interface ISpfxWebPartDemoNjWebPartProps {
  description: string;
}

export default class SpfxWebPartDemoNjWebPart extends BaseClientSideWebPart<ISpfxWebPartDemoNjWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxWebPartDemoNjProps> = React.createElement(
      SpfxWebPartDemoNj,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return super.onInit();
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
