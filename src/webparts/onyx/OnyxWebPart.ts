import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OnyxWebPartStrings';
import Onyx from './components/Onyx';
import { IOnyxProps } from './components/IOnyxProps';

export interface IOnyxWebPartProps {
  adminFlowUrl: string;
  adminEmails: string;
  dataSiteUrl: string;
  description: string;
  staffDirectorySiteUrl: string;
  staffDirectoryListName: string;
}

export default class OnyxWebPart extends BaseClientSideWebPart<IOnyxWebPartProps> {

  public render(): void {
    this.domElement.classList.add('onyxWebPartHost');

    const element: React.ReactElement<IOnyxProps> = React.createElement(
      Onyx,
      {
        adminFlowUrl: this.properties.adminFlowUrl || '',
        adminEmails: this.properties.adminEmails || '',
        context: this.context,
        currentUserEmail: this.context.pageContext.user.email,
        currentUserId: (this.context.pageContext.legacyPageContext as { userId?: number }).userId,
        dataSiteUrl: this.properties.dataSiteUrl || this.context.pageContext.web.absoluteUrl,
        staffDirectorySiteUrl: this.properties.staffDirectorySiteUrl || this.properties.dataSiteUrl || this.context.pageContext.web.absoluteUrl,
        staffDirectoryListName: this.properties.staffDirectoryListName || 'PACT Staff Directory',
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    document.body.classList.add('onyxFullScreenHost');

    if (!this.properties.dataSiteUrl) {
      this.properties.dataSiteUrl = this.context.pageContext.web.absoluteUrl;
    }

    if (!this.properties.staffDirectoryListName) {
      this.properties.staffDirectoryListName = 'PACT Staff Directory';
    }

    if (!this.properties.staffDirectorySiteUrl) {
      this.properties.staffDirectorySiteUrl = this.properties.dataSiteUrl;
    }

    return Promise.resolve();
  }

  protected onDispose(): void {
    document.body.classList.remove('onyxFullScreenHost');
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
                PropertyPaneTextField('dataSiteUrl', {
                  label: 'Product Brief data site URL',
                  description: 'Use the exact site URL where OnyxProductBriefs, OnyxBriefFeatures, and OnyxBriefRoles were created.'
                }),
                PropertyPaneTextField('staffDirectorySiteUrl', {
                  label: 'Staff directory site URL',
                  description: 'Use the exact site URL where PACT Staff Directory exists.'
                }),
                PropertyPaneTextField('staffDirectoryListName', {
                  label: 'Staff directory list name',
                  description: 'Default: PACT Staff Directory. This will be used when Product Brief is wired to staff metadata.'
                }),
                PropertyPaneTextField('adminFlowUrl', {
                  label: 'Admin notification flow URL',
                  description: 'HTTP trigger URL called after a successful submitted brief. Draft saves do not call this flow.',
                  multiline: true
                }),
                PropertyPaneTextField('adminEmails', {
                  label: 'Admin email addresses',
                  description: 'Comma-separated emails that can see the admin dashboard and export all visible records.'
                }),
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
