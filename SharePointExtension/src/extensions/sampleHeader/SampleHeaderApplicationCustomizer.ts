import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { ReactHeader } from './components/ReactHeader';
import { IReactHeaderProps } from './components/IReactHeaderProps';

import * as strings from 'SampleHeaderApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SampleHeaderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISampleHeaderApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SampleHeaderApplicationCustomizer
  extends BaseApplicationCustomizer<ISampleHeaderApplicationCustomizerProperties> {

  private _headerPlaceholder: PlaceholderContent | undefined;

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
    }

    const element: React.ReactElement<IReactHeaderProps> = React.createElement(
      ReactHeader,
      {
        description: "The default header"
      }
    );

    ReactDom.render(element, this._headerPlaceholder.domElement);
  }

  private _onDispose(): void {
    console.log('[CustomHeader._onDispose] Disposed custom header.');
  }
}
