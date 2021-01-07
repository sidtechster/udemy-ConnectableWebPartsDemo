import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ConsumerWebPartDemoWebPartStrings';
import ConsumerWebPartDemo from './components/ConsumerWebPartDemo';
import { IConsumerWebPartDemoProps } from './components/IConsumerWebPartDemoProps';

import {  
  DynamicDataSharedDepth,
  PropertyPaneDynamicField,
  PropertyPaneDynamicFieldSet,
  IPropertyPaneConditionalGroup,
  IWebPartPropertiesMetadata
} from '@microsoft/sp-webpart-base';

import { DynamicProperty } from "@microsoft/sp-component-base";

export interface IConsumerWebPartDemoWebPartProps {
  description: string;
  DeptTitleId: DynamicProperty<string>;
}

export default class ConsumerWebPartDemoWebPart extends BaseClientSideWebPart<IConsumerWebPartDemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IConsumerWebPartDemoProps> = React.createElement(
      ConsumerWebPartDemo,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        DeptTitleId: this.properties.DeptTitleId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'DeptTitleId': { dynamicPropertyType: 'string' }
    };
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
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Select Department ID',
                  fields: [
                    PropertyPaneDynamicField('DeptTitleId', {
                      label: 'Department ID'
                    })
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel: 'Select the web part containing the list of departments'
                    }
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
