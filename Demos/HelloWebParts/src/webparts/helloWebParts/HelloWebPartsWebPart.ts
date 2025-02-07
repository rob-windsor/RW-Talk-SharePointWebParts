import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient } from '@microsoft/sp-http';
import { IListItem } from '../../models/IListItem';

import styles from './HelloWebPartsWebPart.module.scss';
import * as strings from 'HelloWebPartsWebPartStrings';

export interface IHelloWebPartsWebPartProps {
  description: string;
  listName: string;
}

export default class HelloWebPartsWebPart extends BaseClientSideWebPart<IHelloWebPartsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listItems: IListItem[] = [];

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWebParts} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Description property value: <strong>${escape(this.properties.description)}</strong></div>
        <div>List name property value: <strong>${escape(this.properties.listName)}</strong></div>
      </div>
      <div class="${styles.mySection}">
        <button type="button" id="welcomeButton">Show welcome message</button>
      </div>
      <div class="${styles.mySection}">
        <button type="button" id="getItemsButton">Get List Items</button>
      </div>
      <div class="${styles.mySection}">
        <ul>
          ${this._listItems.map(item => 
            `<li><strong>Id:</strong> ${item.Id}, <strong>Title:</strong> ${item.Title}</li>`)
            .join('')}
        </ul>
      </div>
    </section>`;

    const welcomeButton = this.domElement.querySelector('#welcomeButton');
    if (welcomeButton) {
      welcomeButton.addEventListener('click', (event: MouseEvent) => {
        event.preventDefault();
        alert('Welcome to the SharePoint Framework!');
      });
    }

    const getItemsButton = this.domElement.querySelector('#getItemsButton');
    if (getItemsButton) {
      getItemsButton.addEventListener('click', async (event: MouseEvent) => {
        const response = await this.context.spHttpClient.get(
          this.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('${this.properties.listName}')/items?$select=Id,Title`,
          SPHttpClient.configurations.v1);
      
        if (!response.ok) {
          const responseText = await response.text();
          throw new Error(responseText);
        }
      
        const responseJson = await response.json();
    
        this._listItems = responseJson.value as IListItem[];
        this.render();
      });
    }
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
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
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
    const {
      semanticColors
    } = currentTheme;

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
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
