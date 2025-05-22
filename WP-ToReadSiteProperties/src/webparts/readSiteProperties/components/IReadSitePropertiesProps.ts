import { EnvironmentType } from "@microsoft/sp-core-library";

export interface IReadSitePropertiesProps {
  environment: EnvironmentType;
  environemtTitle: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  absoluteUrl: string
  siteTitle: string;
  relativeUrl: string;
}
