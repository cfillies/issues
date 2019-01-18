import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  // This variable has been added
  private _teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    var teamName: string;

    if (this._teamsContext) {
      teamName = this._teamsContext.teamName;
    } else {
      teamName = "NOT IN TEAMS.";
    }

    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        teamsTitle: teamName,
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
