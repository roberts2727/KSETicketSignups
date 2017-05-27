import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './KseTicketSignups.module.scss';
import * as strings from 'kseTicketSignupsStrings';
import { IKseTicketSignupsWebPartProps } from './IKseTicketSignupsWebPartProps';

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import pnp from "sp-pnp-js";
import { Item, ItemUpdateResult } from '../../../node_modules/sp-pnp-js/lib/sharepoint/items';
import { ISPList } from './KseTicketSignupsWebPart';
export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: number;
    Day: string;
    jcaa: Date;
    Register:{Description: string, Url: string};
    Alloted: number;
    Remaining: number;
    
  }
  export default class KseTicketSignupsWebPart extends BaseClientSideWebPart<IKseTicketSignupsWebPartProps> {
   
 
 private _getListItemData(): Promise<ISPLists> {
   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('Games')/items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
   });
 }
 private _renderList(items: ISPList[]): void {
   
    let html: string = '';
    items.forEach((item: ISPList) => {
      
      
      if(item.Remaining>0){
      html += `
        <ul class="${styles.list}">
            <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}<br>${item.Day}<br>${item.jcaa}<br>Tickets Allotted: ${item.Alloted}<br>Tickets Remaining: ${item.Remaining}<br>
                  <button class="${styles.button} update-Button">
            <span class="${styles.label}">Register!</span>
          </button>
                </span>
            </li>
        </ul>`;
      } 
        else{
      html += `
        <ul class="${styles.list}">
            <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}<br>${item.Day}<br>${item.jcaa}<br>Tickets Allotted: ${item.Alloted}<br>Tickets Remaining: ${item.Remaining}<br>Sorry, Game is Closed.</span>
            </li>
        </ul>`;}
                    });
       const listContainer: Element = this.domElement.querySelector('#spListContainer');
       listContainer.innerHTML = html;
      }
  public render(): void {
     this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to KSE Ticket Signups!</span>
              <p class="ms-font-l ms-fontColor-white">Choose a Game from the list of upcoming games to register.</p>
              <p class="ms-font-l ms-fontColor-white">Loading from ${escape(this.context.pageContext.web.title)}</p>
              <a href="https://ksedev.sharepoint.com/sites/dev1/CDN/TicketPolicy.docx?d=w0f15f5b6f2a04939bd9085c694ea0bc1" class="${styles.button}">
                <span class="${styles.label}">Read Comp Ticket Policy</span>
              </a>
            </div>
          </div>
        </div>  
        <div id="spListContainer" />
      </div>`;
     this._getListItemData()
          .then((response) => {
          this._renderList(response.value);
          console.log(response.value);
          this.setButtonsEventHandlers();
          this.setButtonsState();
                         
        });
    }
    
  private setButtonsState(): void {
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll(`button.${styles.button}`);
    const listNotConfigured: boolean = this.listNotConfigured();

    for (let i: number = 0; i < buttons.length; i++) {
      const button: Element = buttons.item(i);
      if (listNotConfigured) {
        button.setAttribute('disabled', 'disabled');
      }
      else {
        button.removeAttribute('disabled');
      }
    }
  }
  
  private setButtonsEventHandlers(): void {
    const webPart: KseTicketSignupsWebPart = this;
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll(`button.${styles.button}`);
    
    for (let i: number = 0; i < buttons.length; i++) {
      const button: Element = buttons.item(i);
      button.addEventListener('click', () => { webPart.updateItem(); });
  }}
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
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
    
  }
   private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      pnp.sp.web.lists.getByTitle(this.properties.listName)
        .items.orderBy('Id', true).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  
  private updateItem(): void {
   
    let latestItemId: number = undefined;
    let etag: string = undefined;

    this.getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
          return pnp.sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(itemId).get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<ISPList> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as ISPList);
      })
      .then((item: ISPList): Promise<ItemUpdateResult> => {
        return pnp.sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(item.Id).update({
            'Remaining' : `${item.Remaining - 1}`
          }, etag);
      })
      .then((result: ItemUpdateResult): void => {
        console.log(`Item with ID: ${latestItemId} successfully updated`);
      }, (error: any): void => {
        console.log('Loading latest item failed with error: ' + error);
      });
  }
  private listNotConfigured(): boolean {
    return this.properties.listName === undefined ||
      this.properties.listName === null ||
      this.properties.listName.length === 0;
  }
}
