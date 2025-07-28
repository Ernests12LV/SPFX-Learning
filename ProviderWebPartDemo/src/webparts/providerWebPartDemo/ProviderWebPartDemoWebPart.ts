import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import * as strings from 'ProviderWebPartDemoWebPartStrings';
import ProviderWebPartDemo from './components/ProviderWebPartDemo';
import { IProviderWebPartDemoProps } from './components/IProviderWebPartDemoProps';
import { ISharedData } from './components/ISharedData';
import { DataProvider } from './components/DataProvider';

export interface IProviderWebPartDemoWebPartProps {
  description: string;
}

export default class ProviderWebPartDemoWebPart extends BaseClientSideWebPart<IProviderWebPartDemoWebPartProps> implements IDynamicDataCallables {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _dataProvider: DataProvider;

  protected onInit(): Promise<void> {
    // Initialize the data provider
    const initialData: ISharedData = {
      message: 'Hello from Provider Web Part!',
      timestamp: new Date(),
      userInfo: {
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email || 'N/A'
      },
      counter: 0
    };

    this._dataProvider = new DataProvider(initialData);

    // Register this web part as a dynamic data source
    this.context.dynamicDataSourceManager.initializeSource(this);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  public render(): void {
    const element: React.ReactElement<IProviderWebPartDemoProps> = React.createElement(
      ProviderWebPartDemo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        dataProvider: this._dataProvider,
        onDataUpdate: this._onDataUpdate.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onDataUpdate(updates: Partial<ISharedData>): void {
    this._dataProvider.updateData(updates);
    // Notify dynamic data consumers that data has changed
    this.context.dynamicDataSourceManager.notifyPropertyChanged('sharedData');
  }

  /**
   * Define the dynamic data properties that this web part provides
   */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'sharedData',
        title: 'Shared Data from Provider'
      },
      {
        id: 'message',
        title: 'Message'
      },
      {
        id: 'counter',
        title: 'Counter Value'
      },
      {
        id: 'userInfo',
        title: 'User Information'
      }
    ];
  }

  /**
   * Return the current value of the specified dynamic data property
   */
  public getPropertyValue(propertyId: string): any {
    switch (propertyId) {
      case 'sharedData':
        return this._dataProvider.data;
      case 'message':
        return this._dataProvider.data.message;
      case 'counter':
        return this._dataProvider.data.counter;
      case 'userInfo':
        return this._dataProvider.data.userInfo;
      default:
        throw new Error('Bad property id');
    }
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
