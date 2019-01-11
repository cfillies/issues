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

import { taxonomy, StringMatchOption } from '@pnp/sp-taxonomy';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  protected async onInit(): Promise<void> {
    taxonomy.setup({
      spfxContext: this.context
    });

    const termStore = await taxonomy.getDefaultSiteCollectionTermStore().usingCaching().get();

    const terms = await termStore.getTerms({
      TermLabel: 'Dev',
      StringMatchOption: StringMatchOption.StartsWith,
      DefaultLabelOnly: true,
      TrimUnavailable: true,
      ResultCollectionSize: 10
    }).usingCaching().get();
    console.log(terms);

    let batch = taxonomy.createBatch();

    terms.forEach(term => {
      term.termSet.inBatch(batch).usingCaching().get().then(termSet => {
        console.log(termSet.Id);
      });
      term.labels.inBatch(batch).usingCaching().get().then(labels => {
        console.log(labels);
      });
    });

    batch.execute().then(() => {
      console.log('executed');
    }, (error) => {
      console.log(error);
    });


    return super.onInit();
  }


  public render(): void {
    const element: React.ReactElement<IHelloWorldProps > = React.createElement(
      HelloWorld,
      {
        description: this.properties.description
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
