import * as React from "react";
import * as ReactDom from "react-dom";
import { Environment, EnvironmentType, Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "Tempalte120WebPartStrings";
import Tempalte120 from "./components/Tempalte120";
import { ITempalte120Props } from "./components/ITempalte120Props";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ITempalte120WebPartProps {
  description: string;
  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  isCertified: boolean;
  category: string;
  deliveryOption: string;
  features: string[];
  paymentMethod: string;
  colorScheme: string;
}

export interface ISharePointList {
  Title: string;
  Id: string;
}

export interface ISharePointLists {
  value: ISharePointList[];
}

export default class Tempalte120WebPart extends BaseClientSideWebPart<ITempalte120WebPartProps> {
  private _lists: ISharePointList[] = [];

  private _getListsOfLists(): Promise<ISharePointLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getAndRenderLists(): Promise<void> {
    if (Environment.type === EnvironmentType.Local) {
      // Handle local environment if needed
      return Promise.resolve();
    } else {
      return this._getListsOfLists().then((response) => {
        this._lists = response.value;
      });
    }
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  private calculateBillAmount(): number {
    const cost = this.properties.productCost || 0;
    const quantity = this.properties.quantity || 0;
    return cost * quantity;
  }

  private calculateDiscount(billAmount: number): number {
    // Example: 10% discount for bills over 1000
    return billAmount > 1000 ? billAmount * 0.1 : 0;
  }

  private calculateNetBillAmount(billAmount: number, discount: number): number {
    return billAmount - discount;
  }

  public render(): void {
    const billAmount = this.calculateBillAmount();
    const discount = this.calculateDiscount(billAmount);
    const netBillAmount = this.calculateNetBillAmount(billAmount, discount);

    const element: React.ReactElement<ITempalte120Props> = React.createElement(
      Tempalte120,
      {
        productName: this.properties.productName || '',
        productDescription: this.properties.productDescription || '',
        productCost: this.properties.productCost || 0,
        quantity: this.properties.quantity || 0,
        billAmount: billAmount,
        discount: discount,
        netBillAmount: netBillAmount,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        lists: this._lists,
        isCertified: this.properties.isCertified,
        rating: (this.properties as any).rating || 0,
        category: this.properties.category || 'electronics',
        deliveryOption: this.properties.deliveryOption || '',
        features: this.properties.features || [],
        paymentMethod: this.properties.paymentMethod || '',
        colorScheme: this.properties.colorScheme || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.properties.productName = 'initial product name';
    this.properties.productDescription = 'initial product description';
    this.properties.productCost = 69;
    this.properties.quantity = 69;

    return Promise.all([
      this._getEnvironmentMessage().then((message) => {
        this._environmentMessage = message;
      }),
      this._getAndRenderLists()
    ]).then(() => {
      return Promise.resolve();
    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
              groupName: "Product Details",
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                }),
                PropertyPaneTextField('productDescription', {
                  label: "Product Description",
                  multiline: true
                }),
                PropertyPaneTextField('productCost', {
                  label: "Product Cost ($)",
                  onGetErrorMessage: this.validateNumber
                }),
                PropertyPaneTextField('quantity', {
                  label: "Quantity",
                  onGetErrorMessage: this.validateNumber
                }),
                PropertyPaneToggle('isCertified', {
                  label: "Is Product Certified",
                  onText: "Yes",
                  offText: "No"
                }),
                PropertyPaneSlider('rating', {
                  label: "Product Rating",
                  min: 0,
                  max: 10,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneChoiceGroup('category', {
                  label: "Product Category",
                  options: [
                    { key: 'electronics', text: 'Electronics' },
                    { key: 'clothing', text: 'Clothing' },
                    { key: 'books', text: 'Books' }
                  ]
                })
              ]
            },
            {
              groupName: "Additional Options",
              groupFields: [
                PropertyPaneChoiceGroup('deliveryOption', {
                  label: "Delivery Method",
                  options: [
                    {
                      key: 'standard',
                      text: 'Standard Delivery',
                      iconProps: {
                        officeFabricIconFontName: 'Mail'
                      }
                    },
                    {
                      key: 'express',
                      text: 'Express Delivery',
                      iconProps: {
                        officeFabricIconFontName: 'MailAlert'
                      }
                    }
                  ]
                }),
                PropertyPaneDropdown('paymentMethod', {
                  label: "Payment Method",
                  options: [
                    { key: 'credit', text: 'Credit Card' },
                    { key: 'debit', text: 'Debit Card' },
                    { key: 'paypal', text: 'PayPal' },
                    { key: 'crypto', text: 'Cryptocurrency' }
                  ]
                }),
                PropertyPaneChoiceGroup('colorScheme', {
                  label: "Color Scheme",
                  options: [
                    { key: 'light', text: 'Light' },
                    { key: 'dark', text: 'Dark' },
                    { key: 'custom', text: 'Custom' }
                  ]
                }),
                PropertyPaneCheckbox('features', {
                  text: 'Add Insurance',
                  checked: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private validateNumber(value: string): string {
    if (value === null || value.trim().length === 0) return '';
    const number = Number(value);
    return isNaN(number) || number < 0 ? 'Please enter a valid positive number' : '';
  }
}
