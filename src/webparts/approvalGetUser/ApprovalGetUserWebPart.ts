import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import styles from './ApprovalGetUserWebPart.module.scss';
import * as strings from 'ApprovalGetUserWebPartStrings';

export interface IApprovalGetUserWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id:string;
  Description: string;
  Reason: string;
  EmpName: string;
  StartDate: String;
  EndDate: string;
  Status: string;
  Manager:string;
}
export default class ApprovalGetUserWebPart extends BaseClientSideWebPart<IApprovalGetUserWebPartProps> {

  private mainBody:string;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.approvalGetUser}">
        <div class="${ styles.container}">
        <div class="ms-Grid-row ms-bgColor-themeDark $ms-color-themeDark ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
       
        <p class="${styles.title}">Leave Request List</p>
        </div>
      </div>
      <div class="ms-Grid-row ms-bgColor-themeDark $ms-color-themeDark ${styles.row}">
      <br>
      <div id="spListContainer" />
      </div>
    </div>
  </div>`;

    this._renderListAsync();
  }

  private _renderListAsync(): void {

    let items:ISPList[]=[];

    this._getListData()
      .then((response) => {       
        let itemsTemp:ISPList[]=response.value;
        console.log(itemsTemp);
        var i=0;
        itemsTemp.forEach(element => {
          if(element.Manager===this.context.pageContext.user.email){
            items[i++]=element;
          }
        });
        console.log(items);
        this._renderList(items);
      });
  }

  private _getListData(): Promise<ISPLists> {
    //console.log(this.context.pageContext.web.absoluteUrl)
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('EmployeeApproval')/Items`,
      SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        //  console.log(response.json());
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>Description</th><th>Name</th><th>Reason</th><th>From Date</th><th>To Date</th><th>Status</th><th>Approve/Reject</th>`;
    items.forEach((item: ISPList) => {
      if(item.EmpName !=this.context.pageContext.user.displayName){
      html += `
          <tr>
              <td>${item.Description}</td>
              <td>${item.EmpName}</td>
              <td>${item.Reason}</td>
              <td>${item.StartDate} </td>
              <td>${item.EndDate} </td>
              <td>${item.Status} </td>`;
              html+=`<td>`;
              if(item.Status == 'Pending'){
              
                html+=`<button name="Apbtn" id=${item.Id} style="
                background-color: #008CBA;
                padding: 7PX 23PX;
                COLOR: WHITE;
                FONT-SIZE: 11PX;
                CURSOR: POINTER;
                MARGIN: 4PX 2PX;
                BORDER: NONE; class="Aprbtn">Approve</button>

                <button name="Rebtn" id=${item.Id} style="
                background-color: #008CBA;
                padding: 7PX 23PX;
                COLOR: WHITE;
                FONT-SIZE: 11PX;
                CURSOR: POINTER;
                MARGIN: 4PX 2PX;
                BORDER: NONE;class="Rejbtn">Reject</button>`;
              }
              else{
                html+=`<p style="color:green;">No action required</p>`;
              }
           html+=`</td></tr>`;
            }
    });
    html += `</table>`;

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    const events: ApprovalGetUserWebPart = this;
    events.fetchSubmissionId();
  }

  protected fetchSubmissionId() {
    const events: ApprovalGetUserWebPart = this;
    //To gte the array of approve buttons
    var btnsApp = document.getElementsByName("Apbtn");
    for (let index = 0; index < btnsApp.length; index++) {
      btnsApp[index].addEventListener("click", () => {
        events.onStatusChange(btnsApp[index].getAttribute("id"), true);
      });
    }
    //To gte the array of reject buttons
    var btnsRej = document.getElementsByName("Rebtn");
    for (let index = 0; index < btnsRej.length; index++) {
      btnsRej[index].addEventListener("click", () => {
        events.onStatusChange(btnsRej[index].getAttribute("id"), false);
      });
    }
  }

  protected onStatusChange(submissionId: string, isApproved: boolean): void {

    //to get the specified list item only
    this.getItemById(submissionId).then((response: SPHttpClientResponse): Promise<ISPList> => {
      return response.json();
    }).then((item: ISPList): void => {
      console.log(item.Id + " , status: " + item.Status);
      if (isApproved) {
        const body: string = JSON.stringify({
          'Status': 'Approved'
        });
        this.mainBody=body;
      }
      else {
        const body: string = JSON.stringify({
          'Status': 'Rejected'
        });
        this.mainBody=body;
      }
      //functiopn to update it
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeApproval')/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: this.mainBody
        })
        .then((response: SPHttpClientResponse): void => {
          console.log("successfully updated");
        }, (error: any): void => {
          console.log(`Error updating item: ${error}`);
        });
    });
    location.reload(true);
  }

  private getItemById(submissionId: string): Promise<SPHttpClientResponse> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('EmployeeApproval')/items(${submissionId})?$select=Status,ID`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      });
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
