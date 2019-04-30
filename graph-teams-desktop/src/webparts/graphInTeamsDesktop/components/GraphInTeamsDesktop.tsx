import * as React from 'react';
import styles from './GraphInTeamsDesktop.module.scss';
import { IGraphInTeamsDesktopProps } from './IGraphInTeamsDesktopProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { User } from '@microsoft/microsoft-graph-types';
import { List } from 'office-ui-fabric-react';

export interface IGraphInTeamsDesktopState {
  teamMembers: User[];
}
export default class GraphInTeamsDesktop extends React.Component<IGraphInTeamsDesktopProps, IGraphInTeamsDesktopState> {
  private _teamsContext: any;
  private isteams: boolean = false;
  constructor(props: IGraphInTeamsDesktopProps) {
    super(props);
    this.state = {
      teamMembers: []
    };
  }
  public componentDidMount() {
    this.isteams = (this.props.context.microsoftTeams != null && this.props.context.microsoftTeams != undefined);
    if (this.isteams) {
      let _retVal: Promise<any> = Promise.resolve();
      _retVal = new Promise((resolve, _reject) => {
        if (this.props.context.microsoftTeams != undefined) {
          this.props.context.microsoftTeams.getContext(context => {
            let teamid: string | undefined = undefined;
            this._teamsContext = context;
            teamid = this._teamsContext.groupId;

            if (teamid != undefined) {
              this.props.graphClient.api(`groups/${teamid}/members`).version('v1.0').get().then(members => {
                console.log(members);
                this.setState({ teamMembers: members });
              }).catch(error => {
                console.log(error);
                this.setState({ teamMembers: [] });
              });
            }
            resolve();
          });
        }
      });
      console.log(_retVal);
    }
  }
  public render(): React.ReactElement<IGraphInTeamsDesktopProps> {
    return (
      <div className={styles.graphInTeamsDesktop}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
              { this.isteams &&
                <List items={this.state.teamMembers} onRenderCell={this._onRenderCell}></List>
              }
            </div>
          </div>
        </div>
      </div>
    );
  }
  private _onRenderCell(item: User, index: number | undefined): JSX.Element {
    return (
      <div>{item.displayName}</div>
    );
  }
}
