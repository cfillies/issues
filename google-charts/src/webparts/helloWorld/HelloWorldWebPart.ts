import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import * as google from 'google';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  protected onInit(): Promise<void> {
    google.charts.load("current", { packages: ["corechart"] });
    google.charts.setOnLoadCallback(this._drawChart.bind(this));

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld}" id="pie-chart">
      </div>`;
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

  private _drawChart() {
    const data = google.visualization.arrayToDataTable([ 
      ['Task', 'Hours per Day'], 
      ['Work', 11], 
      ['Eat', 2], 
      ['Commute', 2], 
      ['Watch TV', 2], 
      ['Sleep', 7] ]);
    const options = {
      title: 'My Daily Activities',
      pieHole: 0.4,
    };

    const chart = new google.visualization.PieChart(document.getElementById('pie-chart'));
    chart.draw(data, options);
  }
}
