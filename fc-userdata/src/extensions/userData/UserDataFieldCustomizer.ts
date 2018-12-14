import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log, Environment } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'UserDataFieldCustomizerStrings';
import UserData, { IUserDataProps } from './components/UserData';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IUserDataFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'UserDataFieldCustomizer';

export default class UserDataFieldCustomizer
  extends BaseFieldCustomizer<IUserDataFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated UserDataFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "UserDataFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {

    // userData check...
    if (!event.userData) {
      console.log(`onRenderCell. ItemId: ${event.listItem.getValueByName('ID')}. UserData is undefined`);
      event.userData = {
        counter: 1
      };
    }
    else {
      console.log(`onRenderCell. ItemId: ${event.listItem.getValueByName('ID')}. UserData.counter: ${event.userData.counter}`);
    }


    // Use this method to perform your custom cell rendering.
    const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;

    const userData: React.ReactElement<{}> =
      React.createElement(UserData, { text } as IUserDataProps);

    ReactDOM.render(userData, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    console.log(`onDisposeCell. UserData:  ${event.userData}`);
    delete event.userData;
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
