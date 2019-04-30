import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IGraphInTeamsDesktopProps {
  description: string;
  context: WebPartContext;
  graphClient: MSGraphClient;
}
