import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldBlankWebPartStrings';
import HelloWorldBlank from './components/HelloWorldBlank';
import { IHelloWorldBlankProps } from './components/IHelloWorldBlankProps';

export interface IHelloWorldBlankWebPartProps {
  description: string;
}

export default class HelloWorldBlankWebPart extends BaseClientSideWebPart<IHelloWorldBlankWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloWorldBlankProps > = React.createElement(
      HelloWorldBlank,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
