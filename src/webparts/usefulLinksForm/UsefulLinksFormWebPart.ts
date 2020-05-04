import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UsefulLinksFormWebPartStrings';
import UsefulLinksForm from './components/UsefulLinksForm';
import { IUsefulLinksFormProps } from './components/IUsefulLinksFormProps';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IUsefulLinksFormWebPartProps {
  description: string;
  SiteUrl: string;
  context: WebPartContext;
}

export default class UsefulLinksFormWebPart extends BaseClientSideWebPart <IUsefulLinksFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUsefulLinksFormProps> = React.createElement(
      UsefulLinksForm,
      {
        description: this.properties.description,
        SiteUrl: this.context.pageContext.web.absoluteUrl,
        context: this.properties.context
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
