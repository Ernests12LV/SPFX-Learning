import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
//import { escape } from "@microsoft/sp-lodash-subset";

//import styles from "./GraphApiDemoWebPart.module.scss";
import * as strings from "GraphApiDemoWebPartStrings";

//import { MSGraphClient } from "@microsoft/sp-http";
//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export interface IGraphApiDemoWebPartProps {
  description: string;
}

interface IUserInfo {
  displayName?: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  mobilePhone?: string;
  error?: string;
}

class UserInfoComponent {
  public static render(
    domElement: HTMLElement,
    user: IUserInfo,
    description: string
  ): void {
    if (user.error) {
      domElement.innerHTML = `<div>Error fetching user data: ${user.error}</div>`;
      return;
    }

    domElement.innerHTML = `
      <div>
        <p><strong>Description:</strong> ${description}</p>
        <p><strong>Display Name:</strong> ${user.displayName ?? "N/A"}</p>
        <p><strong>Given Name:</strong> ${user.givenName ?? "N/A"}</p>
        <p><strong>Surname:</strong> ${user.surname ?? "N/A"}</p>
        <p><strong>Email ID:</strong> ${user.mail ?? "N/A"}</p>
        <p><strong>Mobile Phone:</strong> ${user.mobilePhone ?? "N/A"}</p>
      </div>
    `;
  }
}

export default class GraphApiDemoWebPart extends BaseClientSideWebPart<IGraphApiDemoWebPartProps> {
  private _userInfo: IUserInfo | null = null;

  public async onInit(): Promise<void> {
    await this._fetchUserInfo();
  }

  public render(): void {
    UserInfoComponent.render(
      this.domElement,
      this._userInfo ?? {},
      this.properties.description
    );
  }

  private async _fetchUserInfo(): Promise<void> {
    try {
      const client = await this.context.msGraphClientFactory.getClient("3");
      const user = await client.api("/me").get();
      this._userInfo = {
        displayName: user.displayName,
        givenName: user.givenName,
        surname: user.surname,
        mail: user.mail,
        mobilePhone: user.mobilePhone,
      };
    } catch (error: any) {
      this._userInfo = { error: error.message || "Unknown error" };
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // this._isDarkTheme = !!currentTheme.isInverted;
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

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
