import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IListNamesWebPartProps {
}

export interface IList {
  Title: string;

}

export default class ListNamesWebPart extends BaseClientSideWebPart<IListNamesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <h3>Below are the list names in SPO Site Collection</h3>
    <table id="listNames">
      <thead>
        <tr>
          <th>List Names in the SharePoint Online Site Collection</th>
  
        </tr>
      </thead>
      <tbody ></tbody>
    </table>
      </div>`;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._getListNames()
      .then((lists: IList[]) => {
        const listNames: string[] = lists.map((list: IList) => {
          return `<tr style="background-color:#BDB76B;color:#ffffff;" width: 450px;     height: 50px;    text-align: center;><td >${list.Title}</td></tr>`;
        });

        const listNamesContainer: Element = this.domElement.querySelector('#listNames tbody');
        listNamesContainer.innerHTML = listNames.join('');
      });
  }

  private _getListNames(): Promise<IList[]> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((lists: any) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return lists.value.map((list: any) => {
          return {
            Title: list.Title

          };
        });
      });
  }
}
