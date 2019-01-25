import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ControlStrings';
import CustomPeoplePickerControlWebpart from './components/CustomPeoplePickerControlWebpart';
import { ICustomPeoplePickerControlWebpartProps } from './components/ICustomPeoplePickerControlWebpartProps';

export interface ICustomPeoplePickerControlWebpartWebPartProps {
  description: string;
}

export default class CustomPeoplePickerControlWebpartWebPart extends BaseClientSideWebPart<ICustomPeoplePickerControlWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICustomPeoplePickerControlWebpartProps> = React.createElement(
      CustomPeoplePickerControlWebpart,
      {
        description: this.properties.description,
        context: this.context
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
