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
      
   this.domElement.querySelector('#spListContainer').innerHTML = 
  items.reduce((html: string, item: ISPList) => {
      let Remaining = `<button id="${item.Id}" button class="${styles.button} update-Button">
                         Register!
                       </button>`;
      if (item.Remaining <= 0) Remaining = 'Sorry, Game is Closed.';
      return html += `<li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}
                  <br>${item.Day}
                  <br>${item.jcaa}
                  <br>Tickets Allotted: ${item.Alloted}
                  <br>Tickets Remaining: ${item.Remaining}
                  <br>${Remaining}
                </span>
            </li>`;
    }, `<ul class="${styles.list}"><!--Items go here-->`) + "</ul>";}

  public render(): void {
    let Name: String = undefined;
    let Tickets: String = undefined;
    let Special: String = undefined;
     this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to KSE Ticket Signups.</span>
              <p class="ms-font-l ms-fontColor-white">Choose a Game from the list of upcoming games to register!</p>
              <p class="ms-font-l ms-fontColor-white">Loading from ${escape(this.context.pageContext.web.title)}</p>
              <a href="https://ksedev.sharepoint.com/sites/dev1/CDN/TicketPolicy.docx?d=w0f15f5b6f2a04939bd9085c694ea0bc1" class="${styles.button}">
                <span class="${styles.label}">Read Comp Ticket Policy</span>
              </a>
              <br>
              <br>Name: <input type="text" name="Name">
              <br>
              <br># of Tickets: <input type="text" name="Tickets">
              <br>
              <br>Special Requests: <input type="text" name="Special">
              <br>
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
    var arr = Array.from(this.domElement.querySelectorAll("button."+styles.button));
        arr.forEach(button=>button.addEventListener(
                 'click', (event) => this.updateItem(button.id)));
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
        .items.orderBy('Id', true).select('Id').get()
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

  
  private updateItem(id): void {
   
   
    let latestItemId: number = id;
    let etag: string = undefined;
    console.log(id);
    

    this.getLatestItemId()
      .then((latestItemId: number): Promise<Item> => {
        if (latestItemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = id;
          return pnp.sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(latestItemId).get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<ISPList> => {
        
        return Promise.resolve((item as any) as ISPList);
      })
      .then((item: ISPList): Promise<ItemUpdateResult> => {
        return pnp.sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(latestItemId).update({
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
