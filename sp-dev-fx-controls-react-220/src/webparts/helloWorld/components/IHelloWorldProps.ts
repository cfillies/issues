import { DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHelloWorldProps {
  description: string;
  teamsTitle: string;
  context: WebPartContext;
}
