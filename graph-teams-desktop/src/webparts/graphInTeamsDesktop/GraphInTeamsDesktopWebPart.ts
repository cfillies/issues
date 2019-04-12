import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'GraphInTeamsDesktopWebPartStrings';
import GraphInTeamsDesktop from './components/GraphInTeamsDesktop';
import { IGraphInTeamsDesktopProps } from './components/IGraphInTeamsDesktopProps';

import * as microsoftTeams from '@microsoft/teams-js';

export interface IGraphInTeamsDesktopWebPartProps {
  description: string;
}

export default class GraphInTeamsDesktopWebPart extends BaseClientSideWebPart<IGraphInTeamsDesktopWebPartProps> {

  private _teamsContext: microsoftTeams.Context;

  public onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve) => {
        this.context.microsoftTeams!.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    this._sendRequest();
    const element: React.ReactElement<IGraphInTeamsDesktopProps > = React.createElement(
      GraphInTeamsDesktop,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async _sendRequest() {
    const client = await this.context.msGraphClientFactory.getClient();
    client.api(`groups/${this._teamsContext!.groupId}/members`).version('v1.0').get().then(members => {
      console.log(members);
    }).catch(error => {
      console.log(error);
    });
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
