import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartPropertiesMetadata  
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WebpartIndexingWebPart.module.scss';
import * as strings from 'WebpartIndexingWebPartStrings';

export interface IWebpartIndexingWebPartProps {
  title: string;
  intro: string;
  image: string;
  url: string;
}

export default class WebpartIndexingWebPartWebPart extends BaseClientSideWebPart<IWebpartIndexingWebPartProps> {

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'title': { isSearchablePlainText: true },
      'intro': { isHtmlString: true },
      'image': { isImageSource: true },
      'url': { isLink: true }
     };
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.webpartIndexing}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Demo webpart with indexed properties</span>
              <p class="ms-font-l ms-fontColor-white">Title: ${escape(this.properties.title)}</p>
              <p class="ms-font-l ms-fontColor-white">Introduction: ${escape(this.properties.intro)}</p>
              <p class="ms-font-l ms-fontColor-white">Image Url: ${escape(this.properties.image)}</p>
              <p class="ms-font-l ms-fontColor-white">Link Url: ${escape(this.properties.url)}</p>
            </div>
          </div>
        </div>
      </div>`;
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('intro', {
                  label: strings.IntroFieldLabel
                }),
                PropertyPaneTextField('image', {
                  label: strings.ImageFieldLabel
                }),
                PropertyPaneTextField('url', {
                  label: strings.UrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
