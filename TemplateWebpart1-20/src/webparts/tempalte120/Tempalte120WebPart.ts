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
  PropertyPaneCheckbox,
  PropertyPaneButton,
  PropertyPaneButtonType
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "Tempalte120WebPartStrings";
import Tempalte120 from "./components/Tempalte120";
import { ITempalte120Props } from "./components/ITempalte120Props";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http"; //ISPHttpClientOptions

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
  listName: string;
  listDescription: string;
  itemTitle: string;
  itemId: string;
  listTitle: string;
  listTeam: string;
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
        colorScheme: this.properties.colorScheme || '',
        listName: this.properties.listName || '',
        listDescription: this.properties.listDescription || '',
        itemTitle: this.properties.itemTitle || '',
        listTitle: this.properties.listTitle || '',
        listTeam: this.properties.listTeam || '',
        itemId: this.properties.itemId || ''
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
            description: "Basic Settings"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Basic Product Information",
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: "Product Name"
                }),
                PropertyPaneTextField('productDescription', {
                  label: "Product Description",
                  multiline: true
                })
              ]
            },
            {
              groupName: "Pricing Details",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('productCost', {
                  label: "Product Cost ($)",
                  onGetErrorMessage: this.validateNumber
                }),
                PropertyPaneTextField('quantity', {
                  label: "Quantity",
                  onGetErrorMessage: this.validateNumber
                })
              ]
            },
            {
              groupName: "Product Classifications",
              isCollapsed: true,
              groupFields: [
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
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Advanced Settings"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Categories & Features",
              isCollapsed: false,
              groupFields: [
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
              groupName: "Shipping & Delivery",
              isCollapsed: true,
              groupFields: [
                PropertyPaneChoiceGroup('deliveryOption', {
                  label: "Delivery Method",
                  options: [
                    {
                      key: 'standard',
                      text: 'Standard Delivery',
                      iconProps: { officeFabricIconFontName: 'Mail' }
                    },
                    {
                      key: 'express',
                      text: 'Express Delivery',
                      iconProps: { officeFabricIconFontName: 'MailAlert' }
                    }
                  ]
                }),
                PropertyPaneCheckbox('features', {
                  text: 'Add Insurance',
                  checked: false
                })
              ]
            },
            {
              groupName: "Payment & Appearance",
              isCollapsed: true,
              groupFields: [
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
                })
              ]
            }
          ],
        },
        {
          header: {
            description: "List Management"
          },
          groups: [
            {
              groupName: "Create New List",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "List Name"
                }),
                PropertyPaneTextField('listDescription', {
                  label: "List Description",
                  multiline: true
                }),
                PropertyPaneButton('createList', {
                  text: "Create List",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName) {
                      alert('Please enter a list name');
                      return;
                    }
                    
                    this._checkIfListExists(this.properties.listName)
                      .then((exists: boolean) => {
                        if (exists) {
                          alert(`List "${this.properties.listName}" already exists!`);
                        } else {
                          this._createList(
                            this.properties.listName,
                            this.properties.listDescription || ''
                          );
                        }
                      });
                  }
                })
              ]
            },
            {
              groupName: "Subbit Item to List",
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: "Name"
                }),
                PropertyPaneTextField('listTeam', {
                  label: "Item Title",
                  multiline: true
                }),
                PropertyPaneButton('submitItem', {
                  text: "Submit",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName || !this.properties.listTitle) {
                      alert('Please enter both list name and item title');
                      return;
                    }

                    this._checkIfListExists(this.properties.listName)
                      .then((exists: boolean) => {
                        if (!exists) {
                          alert(`List "${this.properties.listName}" not found!`);
                        } else {
                          this._addItemToList(
                            this.properties.listName,
                            this.properties.listTitle
                          );
                        }
                      });
                  }
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "List Operations"
          },
          groups: [
            {
              groupName: "Operations",
              groupFields: [
                PropertyPaneTextField('itemId', {
                  label: "Item ID"
                }),
                PropertyPaneButton('getAllItems', {
                  text: "Get All Items",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName) {
                      alert('Please enter a list name');
                      return;
                    }
                    this._getAllItems(this.properties.listName)
                      .then(items => {
                        console.log('All Items:', items);
                        alert(`Found ${items.length} items. Check console for details.`);
                      })
                      .catch(error => alert(error.message));
                  }
                }),
                PropertyPaneButton('getItem', {
                  text: "Get Item by ID",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName || !this.properties.itemId) {
                      alert('Please enter both list name and item ID');
                      return;
                    }
                    this._getItemById(this.properties.listName, parseInt(this.properties.itemId))
                      .then(item => {
                        console.log('Item:', item);
                        alert(`Item found. Check console for details.`);
                      })
                      .catch(error => alert(error.message));
                  }
                }),
                PropertyPaneButton('updateItem', {
                  text: "Update Item",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName || !this.properties.itemId) {
                      alert('Please enter both list name and item ID');
                      return;
                    }
                    const updates = {
                      Title: `Updated Title ${new Date().toISOString()}`
                    };
                    this._updateListItem(this.properties.listName, parseInt(this.properties.itemId), updates)
                      .catch(error => alert(error.message));
                  }
                }),
                PropertyPaneButton('deleteItem', {
                  text: "Delete Item",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    if (!this.properties.listName || !this.properties.itemId) {
                      alert('Please enter both list name and item ID');
                      return;
                    }
                    if (confirm('Are you sure you want to delete this item?')) {
                      this._deleteListItem(this.properties.listName, parseInt(this.properties.itemId))
                        .catch(error => alert(error.message));
                    }
                  }
                })
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

  private _createList(listName: string, listDescription: string): Promise<void> {
    const listUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;

    const body: string = JSON.stringify({
      '__metadata': { 'type': 'SP.List' },
      'BaseTemplate': 100,
      'Title': listName,
      'Description': listDescription,
      'AllowContentTypes': true,
      'ContentTypesEnabled': true
    });

    return this.context.spHttpClient.post(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'Odata-Version': '3.0'
        },
        body: body
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json().then(() => {
            alert(`List "${listName}" created successfully!`);
            return this._getAndRenderLists();
          });
        } else {
          return response.json().then((error) => {
            console.error('Error details:', error);
            alert(`Error creating list: ${error.error.message.value}`);
          });
        }
      })
      .catch((error: any) => {
        console.error('Error:', error);
        alert(`Error: ${error.message}`);
      });
  }

private _addItemToList(listName: string, itemTitle: string): Promise<void> {
  const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items`;

  const body: string = JSON.stringify({
    '__metadata': { 
      'type': `SP.Data.${listName.replace(/\s/g, '_x0020_')}ListItem` 
    },
    'Title': itemTitle
  });

  return this.context.spHttpClient.post(
    endpoint,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-type': 'application/json;odata=verbose',
        'OData-Version': '3.0'
      },
      body: body
    }
  )
  .then((response: SPHttpClientResponse) => {
    if (response.ok) {
      alert(`Item "${itemTitle}" added successfully!`);
      return;
    }
    throw new Error(`Error adding item: ${response.statusText}`);
  })
  .catch((error: any) => {
    console.error('Error:', error);
    alert(`Error: ${error.message}`);
  });
}

  // private _createSubSite(subSiteTitle: string, subSiteUrl: string, subSiteDescriptin: string):void {
    
  //   const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add";

  //   const spHttpClientOptions: ISPHttpClientOptions = {
  //     body:`{
  //       "parameters":{
  //         "@odata.type": "SP.webInfoCreationInformation",
  //         "Title": "${subSiteTitle}",
  //         "Url": "${subSiteUrl}",
  //         "Description": "${subSiteDescriptin}",
  //         "Language": 1033,
  //         "WebTemplate": "STS#0",
  //         "UseUniquePermissions": false
  //       }
  //     }`
  //   };

  //   this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
  //   .then((response: SPHttpClientResponse) => {
  //     if(response.status == 200){
  //       alert("New sub site created!");
  //     }
  //     else {
  //       alert("Error :" + response.status + "-" + response.statusText);
  //     }
  //   })

  // }

  private _checkIfListExists(listName: string): Promise<boolean> {
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + 
      `/_api/web/lists/GetByTitle('${listName}')`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      return response.ok;
    });
  }

  private _getAllItems(listName: string): Promise<any[]> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items`;
    
    return this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose'
        }
      }
    )
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json().then(data => data.d.results);
      } else {
        throw new Error(`Error getting items: ${response.statusText}`);
      }
    });
  }

  private _getItemById(listName: string, itemId: number): Promise<any> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})`;
    
    return this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      }
    )
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json().then(data => data.d);
      } else {
        throw new Error(`Error getting item: ${response.statusText}`);
      }
    });
  }

  private _updateListItem(listName: string, itemId: number, updates: any): Promise<void> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})`;
    
    const body: string = JSON.stringify({
      '__metadata': { 
        'type': `SP.Data.${listName.replace(/\s/g, '_x0020_')}ListItem`
      },
      ...updates
    });

    return this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*'
        },
        body: body
      }
    )
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert(`Item updated successfully!`);
      } else {
        throw new Error(`Error updating item: ${response.statusText}`);
      }
    });
  }

  private _deleteListItem(listName: string, itemId: number): Promise<void> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})`;
    
    return this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*'
        }
      }
    )
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert(`Item deleted successfully!`);
      } else {
        throw new Error(`Error deleting item: ${response.statusText}`);
      }
    });
  }
}
