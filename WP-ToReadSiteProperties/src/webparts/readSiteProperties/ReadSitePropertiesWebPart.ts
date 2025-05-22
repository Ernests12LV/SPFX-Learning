import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReadSitePropertiesWebPartStrings';
import ReadSiteProperties from './components/ReadSiteProperties';
import { IReadSitePropertiesProps } from './components/IReadSitePropertiesProps';

import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

export interface IReadSitePropertiesWebPartProps {
  description: string;
}

export default class ReadSitePropertiesWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  _findOutEnvironment(): string {
    if (Environment.type === EnvironmentType.Local) {
      return strings.AppLocalEnvironmentSharePoint;
    } else if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      return strings.AppSharePointEnvironment;
    }
    return '';
  }


  public render(): void {
    const element: React.ReactElement<IReadSitePropertiesProps> = React.createElement(
      ReadSiteProperties,
      { 
        environment: Environment.type,
        environemtTitle: this._findOutEnvironment(),
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        absoluteUrl: this.context.pageContext.web.absoluteUrl,
        siteTitle: this.context.pageContext.web.title,
        relativeUrl: this.context.pageContext.web.serverRelativeUrl

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          // let environmentMessage = `URL: ${this.context.pageContext.web.absoluteUrl}`;
          let environmentMessage = `User Name: ${context.user?.displayName}`;          
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
