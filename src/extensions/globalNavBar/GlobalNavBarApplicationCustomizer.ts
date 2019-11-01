import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import * as strings from 'GlobalNavBarApplicationCustomizerStrings';
import { SPTermStoreService, ISPTermObject } from '../../services/SPTermStoreService';
import { IGlobalNavBarProps } from '../../components/IGlobalNavBarProps';
import { GlobalNavBar } from '../../components/GlobalNavBar';
import * as React from 'react';
import * as ReactDom from 'react-dom';

const LOG_SOURCE: string = 'GlobalNavBarApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGlobalNavBarApplicationCustomizerProperties {
  topMenuTermSet: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GlobalNavBarApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalNavBarApplicationCustomizerProperties> {
  
  private _topPlaceholder: PlaceholderContent | undefined;
  private _globalNavItems: ISPTermObject[];

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    let termStoreService: SPTermStoreService = new SPTermStoreService(this.context);
    this._globalNavItems = await termStoreService.getGlobalNavItemsAsync(this.properties.topMenuTermSet);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholder);
    return Promise.resolve<void>();
  }

  private _renderPlaceholder(): void {
    console.log("GlobalNavBarApplicationCustomizer._renderPlaceHolders()");
    // Handling the top placeholder
    if(!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {onDispose: this._onDispose});
      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if(this._globalNavItems && this._globalNavItems.length > 0){
        console.log(this._globalNavItems);
        const element: React.ReactElement<IGlobalNavBarProps> = React.createElement(
          GlobalNavBar,
          {
            menuItems: this._globalNavItems
          }
        );
        ReactDom.render(element, this._topPlaceholder.domElement)
      }
    }
  }

  private _onDispose(): void {
    console.log('[GlobalNavBarApplicationCustomizer._onDispose] Disposed custom top placeholder.');
  }
}
