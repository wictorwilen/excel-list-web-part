'use strict';

import { HttpClient } from '@microsoft/sp-client-base';
import { IWebPartContext } from '@microsoft/sp-client-preview';


export interface IListsService {
  getListNames(): Promise<string[]>;
  getListColumns(listName: string): Promise<string[]>;
  getListItems(listName: string, listColumns: string[]): Promise<string[][]>
}

export class MockListsService implements IListsService {
  constructor() {
  }
  public getListNames(): Promise<string[]> {
    return new Promise<string[]>(resolve => {
      setTimeout(() => {
        resolve(['List 1', 'List 2']);
      }, 1000);
    });
  }
  public getListColumns(listName: string): Promise<string[]> {
    return new Promise<string[]>(resolve => {
      setTimeout(() => {
        resolve([`${listName}:Column 1`, `${listName}:Column 2`, `${listName}:Column 3`]);
      }, 1000);
    });
  }

  getListItems(listName: string, listColumns: string[]): Promise<string[][]> {
    return new Promise<string[][]>(resolve => {
      setTimeout(() => {
        resolve(
          [
            ['A1', 'B1', 'C1'],
            ['A2', 'B2', 'C2'],
            ['A3', 'B3', 'C3'],
          ]);
      }, 1000);
    });
  }
}

export class ListsService implements IListsService {
  private _httpClient: HttpClient;
  private _webAbsoluteUrl: string;

  public constructor(webPartContext: IWebPartContext) {
    this._httpClient = webPartContext.httpClient as any; // tslint:disable-line:no-any
    this._webAbsoluteUrl = webPartContext.pageContext.web.absoluteUrl;
  }

  public getListNames(): Promise<string[]> {
    return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/?$select=Title`)
      .then((response: Response) => {
        var arr: string[] = [];
        return response.json().then((data) => {
          data.value.forEach(l => {
            arr.push(l.Title);
          });
          return arr;
        });
      });
  }

  public getListColumns(listName: string): Promise<string[]> {
    return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/GetByTitle('${listName}')/Fields?$Select=StaticName`)
      .then((response: Response) => {
        var arr: string[] = [];
        return response.json().then((data) => {
          data.value.forEach(l => {
            arr.push(l.StaticName);
          });
          return arr;
        });
      });
  }
  getListItems(listName: string, listColumns: string[]): Promise<string[][]> {
    return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/GetByTitle('${listName}')/Items?$Select=${encodeURI(listColumns.join(','))}`)
      .then((response: Response) => {
        var data: string[][] = [[]];
        return response.json().then((rows) => {
          rows.value.forEach(row => {
            var dataRow = [];
            listColumns.forEach(c => {
              dataRow.push(row[c]);
            });
            data.push(dataRow);
          });
          return data;
        });
      });

  }
}