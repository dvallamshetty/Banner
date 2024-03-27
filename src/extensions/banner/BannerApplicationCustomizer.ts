import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';
import styles from './banner.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
//import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'BannerApplicationCustomizerStrings';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PermissionKind } from "@pnp/sp/security";
//import { Web } from "@pnp/sp/webs";
//import { _Items } from '@pnp/sp/items/types';
const LOG_SOURCE: string = 'BannerApplicationCustomizer';

export interface IBannerApplicationCustomizerProperties {
  bannerMessage: string;
}

export default class BannerApplicationCustomizer
  extends BaseApplicationCustomizer<IBannerApplicationCustomizerProperties> {
  private _bannerPlaceholder: PlaceholderContent | undefined;  
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    let displayBanner: boolean = false;
    let _hasfullPermissions = false;
    const sp = spfi().using(SPFx(this.context));
    const spWebB = spfi("https://chartercom.sharepoint.com/").using(SPFx(this.context));
    sp.web.currentUserHasPermissions(PermissionKind.FullMask).then(
      result => {
        console.log(result);
        _hasfullPermissions = result;
      }
    ).catch(e => { console.error(e) });
    spWebB.web.lists.getByTitle("BannerDetails").items.top(1).filter("SourceSiteUrl eq '" + webUrl + "'")().then(bannerFound => {
      const _output: any = bannerFound[0];
      const _bannermessage: string = _output.BannerMessage;

      if (_output.OnlyFullControlUsers === false) {
        displayBanner = true;
      }
      else {
        displayBanner = _hasfullPermissions;
      }      
      if (bannerFound) { this._renderBanner(_bannermessage, displayBanner) }
    }).catch(e => { console.error(e) });

  }

  private _renderBanner(bannerMessage: string, condition: boolean): void {

    console.log(bannerMessage)
    this.properties.bannerMessage = bannerMessage;
    // Handling the banner placeholder
    if (!this._bannerPlaceholder && condition) {
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
    console.log('[BannerApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
