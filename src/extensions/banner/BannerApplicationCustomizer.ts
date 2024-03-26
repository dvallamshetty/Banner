import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';
import styles from './banner.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'BannerApplicationCustomizerStrings';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
//import { Web } from "@pnp/sp/webs";
//import { _Items } from '@pnp/sp/items/types';
const LOG_SOURCE: string = 'BannerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IBannerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  bannerMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class BannerApplicationCustomizer
  extends BaseApplicationCustomizer<IBannerApplicationCustomizerProperties> {
  private _bannerPlaceholder: PlaceholderContent | undefined;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const spWebB = spfi("https://chartercom.sharepoint.com/").using(SPFx(this.context));
    spWebB.web.lists.getByTitle("BannerDetails").items.top(1).filter("SourceSiteUrl eq '" + webUrl + "'")().then(bannerFound => {
      if (bannerFound) { this._renderBanner(bannerFound[0].BannerMessage) }
    }).catch(e => { console.error(e) });

  }

  private _renderBanner(bannerMessage: string):void {

    console.log(bannerMessage)
    this.properties.bannerMessage = bannerMessage;
    // Handling the banner placeholder
    if (!this._bannerPlaceholder) {
      this._bannerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });
      // The extension should not assume that the expected placeholder is available.
      if (!this._bannerPlaceholder) { console.error("The expected placeholder (Top) was not found."); return; }

      if (this.properties) {
        let bannerString: string = this.properties.bannerMessage;
        if (!bannerString) {
          bannerString = "(Top property was not defined.)";
        }

        if (this._bannerPlaceholder.domElement) {
          this._bannerPlaceholder.domElement.innerHTML = `
            <div class="${styles.app}">
              <div class="${styles.top}">
              <!-- <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(bannerString)}
                <i class="ms-Icon ms-Icon--Info" aria-hidden="true">${bannerString}</i> -->
                ${bannerString}
              </div>
            </div>`;
        }
      }
    }

  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
