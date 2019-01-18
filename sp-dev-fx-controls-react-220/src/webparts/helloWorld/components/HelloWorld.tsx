import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/ListItemPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <Label>Enquiry type:</Label>
        <ListItemPicker listId='cff84598-349a-483a-9a7b-e13017a19fad'
          columnInternalName='Title'
          itemLimit={1}
          onSelectedItem={this.onSelectedEnquiryType}
          context={this.props.context} />
      </div>
    );
  }

  private onSelectedEnquiryType() {}
}
