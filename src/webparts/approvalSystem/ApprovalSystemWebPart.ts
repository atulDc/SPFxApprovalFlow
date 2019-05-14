import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'ApprovalSystemWebPartStrings';
import ApprovalSystem from './components/ApprovalSystem';
import { IApprovalSystemProps } from './components/IApprovalSystemProps';

export interface IApprovalSystemWebPartProps {
  description: string;
  desc:string;
  reason:string;
  stDate:string;
  endDate:string;
}

export default class ApprovalSystemWebPart extends BaseClientSideWebPart<IApprovalSystemWebPartProps> {
  
  public render(): void {
    const element: React.ReactElement<IApprovalSystemProps> = React.createElement(
      ApprovalSystem,
      {
        description: this.properties.description,
        pagecontext:this.context,
        spHttpClient:this.context.spHttpClient,
        siteURL:this.context.pageContext.web.absoluteUrl,
        desc:this.properties.desc,
        reason:this.properties.reason,
        currentUserName:this.context.pageContext.user.displayName,
        stDate:this.properties.stDate,
        endDate:this.properties.endDate,
        listName:'EmployeeApproval',
        currentUserEmail:this.context.pageContext.user.email
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
