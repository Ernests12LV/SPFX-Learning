import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  DynamicDataSharedDepth
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, DynamicProperty } from '@microsoft/sp-component-base';

import * as strings from 'ConsumerWebPartDemoWebPartStrings';
import ConsumerWebPartDemo from './components/ConsumerWebPartDemo';
import { IConsumerWebPartDemoProps } from './components/IConsumerWebPartDemoProps';
import { ISharedData } from './components/ISharedData';

export interface IConsumerWebPartDemoWebPartProps {
  description: string;
  // Dynamic data properties
  sharedData: DynamicProperty<ISharedData>;
  message: DynamicProperty<string>;
  counter: DynamicProperty<number>;
  userInfo: DynamicProperty<{ displayName: string; email: string; }>;
}

export default class ConsumerWebPartDemoWebPart extends BaseClientSideWebPart<IConsumerWebPartDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IConsumerWebPartDemoProps> = React.createElement(
      ConsumerWebPartDemo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        // Pass dynamic data properties
        sharedData: this.properties.sharedData,
        message: this.properties.message,
        counter: this.properties.counter,
        userInfo: this.properties.userInfo
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Initialize dynamic properties
    this.properties.sharedData = new DynamicProperty<ISharedData>(this.context.dynamicDataProvider);
    this.properties.message = new DynamicProperty<string>(this.context.dynamicDataProvider);
    this.properties.counter = new DynamicProperty<number>(this.context.dynamicDataProvider);
    this.properties.userInfo = new DynamicProperty<{ displayName: string; email: string; }>(this.context.dynamicDataProvider);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    // Register for data change notifications
    this.context.dynamicDataProvider.registerAvailableSourcesChanged(this.render.bind(this));
  }

  protected onDispose(): void {
    // Unregister from data change notifications
    this.context.dynamicDataProvider.unregisterAvailableSourcesChanged(this.render.bind(this));
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
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
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
            },
            {
              groupName: 'Dynamic Data Connection',
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Connect to Provider Data',
                  fields: [
                    PropertyPaneDynamicField('sharedData', {
                      label: 'Shared Data (Complete Object)'
                    }),
                    PropertyPaneDynamicField('message', {
                      label: 'Message Only'
                    }),
                    PropertyPaneDynamicField('counter', {
                      label: 'Counter Only'
                    }),
                    PropertyPaneDynamicField('userInfo', {
                      label: 'User Info Only'
                    })
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
