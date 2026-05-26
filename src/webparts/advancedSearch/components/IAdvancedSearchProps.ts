import { BaseWebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdvancedSearchProps {
    description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any;
  siteUrl: string;
}
