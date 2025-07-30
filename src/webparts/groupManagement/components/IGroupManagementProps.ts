import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IGroupManagementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
   spcontext: WebPartContext;
}
