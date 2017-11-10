//based on: https://github.com/SharePoint/sp-dev-fx-extensions/tree/master/samples/js-application-graph-client
//updated to SPFx 1.3

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'GraphAppCustomApplicationCustomizerStrings';
import styles from './GraphAppCustomApplicationCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset'; 
import { GraphHttpClient, HttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'GraphAppCustomApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGraphAppCustomApplicationCustomizerProperties {
  Header: string;
  Footer: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GraphAppCustomApplicationCustomizer
  extends BaseApplicationCustomizer<IGraphAppCustomApplicationCustomizerProperties> {

  private _headerPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();
    return Promise.resolve<void>();
  }

  @override
  private _renderPlaceHolders(): void {

    // Handling header place holder
    if (!this._headerPlaceholder) {
      this._headerPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._headerPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      // Get group id from page
      let groupId = this.context.pageContext.legacyPageContext.groupId;
      
      if(groupId) {
        // Get group data from graph via new GraphHttpCLient
        this.context.graphHttpClient.get(`v1.0/groups/${groupId}/`, GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
            if (response.ok) {
              return response.json();
            } else {
              console.warn(response.statusText);
            }
          }).then((result: any) => {
            // Set headerstring to the groups display name
            let headerString: string = result.displayName;
            let emailString: string = result.mail;
            let descriptionString: string = result.description;
            
            console.log("Graph API Response");
            console.log(result);

            if (!headerString) {
                headerString = '(Header property was not defined.)';
              }
              if (this._headerPlaceholder.domElement) {
                this._headerPlaceholder.domElement.innerHTML = `
                      <div class="${styles.app}">
                        <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.header}">
                          <b>${escape(headerString)}</b>&nbsp;<i>(${escape(emailString)})</i>: 
                          ${escape(descriptionString)}
                        </div>
                      </div>`;
              }
          });
        
      }
      else
      {
        this._headerPlaceholder.domElement.innerHTML = `
                      <div class="${styles.app}">
                        <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.header}">
                          <p>Group Id not available. This sample only works on Group sites!</p>
                        </div>
                      </div>`;
      }
     
    }
  }

   private _onDispose(): void {
    console.log('[CustomHeader._onDispose] Disposed custom header.');
  }
}
