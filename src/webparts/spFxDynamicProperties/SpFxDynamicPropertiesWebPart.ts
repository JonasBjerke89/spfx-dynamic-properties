import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-client-preview';

import styles from './SpFxDynamicProperties.module.scss';
import * as strings from 'spFxDynamicPropertiesStrings';
import { ISpFxDynamicPropertiesWebPartProps } from './ISpFxDynamicPropertiesWebPartProps';
import MockHttpClient from './MockHttpClient';
import { EnvironmentType } from '@microsoft/sp-client-base';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class SpFxDynamicPropertiesWebPart extends BaseClientSideWebPart<ISpFxDynamicPropertiesWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  private init(): void {
    if (this.context.environment.type === EnvironmentType.Local) {
    this._getMockListData().then((response) => {
      this.lists = response.value.map((list: ISPList) => {
        return {
          key: list.Id,
          text: list.Title
        };
      });
    }); }
    else {
    this._getListData()
      .then((response) => {
        this.lists = response.value.map((list: ISPList) => {
        return {
          key: list.Id,
          text: list.Title
        };
      });
      });
    }
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
        const listData: ISPLists = {
            value:
            [
                { Title: 'Mock List One', Id: '1' },
                { Title: 'Mock List Two', Id: '2' },
                { Title: 'Mock List Three', Id: '3' }
            ]
            };

        return listData;
    }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
  return this.context.httpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`)
      .then((response: Response) => {
      return response.json();
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spFxDynamicProperties}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
              <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
                <span class="ms-Button-label">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

    this.init();
  }

  private lists: IPropertyPaneDropdownOption[] = [];
  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneDropdown('list', {
                  label: 'Choose list',
                  options: this.lists
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
