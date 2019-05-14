import * as React from 'react';
import styles from './ApprovalSystem.module.scss';
import { IApprovalSystemProps } from './IApprovalSystemProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class ApprovalSystem extends React.Component<IApprovalSystemProps, {}> {

  public getEmail:any;
  private desc:string;
  private name:string;
  private reason:string;
  private stDate:string;
  private endDate:string;
  public life:string;

  public render(): React.ReactElement<IApprovalSystemProps> {
    this.isAllFilled();
    var dt=new Date();
    var dateRestrict=dt.getFullYear().toString()+"-"+dt.getMonth().toString()+"-"+dt.getDate().toString();
    return (
      <div className={styles.approvalSystem}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Employee Leave Application</span>
              <p className={styles.subTitle}>Please fill all the details.</p>
              {/* <p className={ styles.description }>{escape(this.props.description)}</p> */}
              <table>
                <tr>
                  <td><label className={styles.subTitle}>Description</label></td>
                  <td><input required  className={styles.detailInput}  type="text" id="desc" placeholder="Description" value={this.props.desc} onChange={this.handleDescChange} />  </td>
                </tr>
                <br />
                <tr>
                  <td><label className={styles.subTitle}>Reason</label></td>
                  <td><input required className={styles.detailInput} type="text" id="reason" placeholder="Reason" value={this.props.reason} onChange={this.handleReasonChange} /></td>
                </tr>
                {/* <br />
                <tr>
                  <td><label className={styles.subTitle}>Name</label></td>
                  <td><input required className={styles.detailInput} type="text" id="name" placeholder="Name" value={this.props.name} onChange={this.handleNameChange}/>  </td>
                </tr> */}
                <br />
                <tr>
                  <td><label className={styles.subTitle}>Start Date</label></td>
                  <td>
                  <input type="date" id="stdate" value={this.props.stDate} onChange={this.handleStartDateChange} />
                  </td>
                </tr>
                <br />
                <tr>
                  <td><label className={styles.subTitle}>End Date</label></td>
                  <td>
                    <input type="date" id="endDate" value={this.props.endDate} onChange={this.handleEndDateChange} />
                  </td>
                </tr>
                <tr>
                  <td><label className={styles.subTitle}>Manager</label>&nbsp</td>
                  <td>
                    <PeoplePicker 
                      context={this.props.pagecontext}
                      titleText="People Picker"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users
                      showtooltip={true}
                      isRequired={true}
                      disabled={false}
                      selectedItems={this._getPeoplePickerItems}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={2000} />
                  </td>
                </tr>
                <br />
                <br />
              </table>
              <div className={styles.wrapButton}>

                <button className={styles.button}  onClick={() => this.createItem()}>
                  <span className={styles.label}>Submit</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private isAllFilled():boolean{
    console.log(this.stDate);
    console.log(this.endDate);
    console.log(this.getEmail);
    if(this.stDate !== null && this.endDate !== null && localStorage.getItem("ManagerEmail") !== null && this.stDate !== undefined && this.endDate !== undefined && localStorage.getItem("ManagerEmail") !== undefined)
    return true;
    else
    return false;
  }

  private handleDescChange = (event) => {
    this.desc=event.target.value;

  }
  private handleReasonChange = (event) => {
    this.reason=event.target.value;

  }
  private handleNameChange = (event) => {
    this.name=event.target.value;
  }

  private handleStartDateChange = (event) => {
    this.stDate=event.target.value;
  }

  private handleEndDateChange = (event) => {
    this.endDate=event.target.value; 
    
  }

  private _getPeoplePickerItems(items: any[]) {
    let managerTemp;
    console.log('Items:', items);
    items.forEach(element => {
       managerTemp=element;
    });
    var getEmailTemp='example@domain.com';
    console.log("Hiiii: "+getEmailTemp);
    console.log("Manager Email:"+managerTemp['secondaryText']);      
    getEmailTemp=managerTemp['secondaryText'];
    
    console.log("Manager Email2:"+getEmailTemp);
    localStorage.setItem("ManagerEmail",getEmailTemp);
    localStorage.setItem("ManagerName",managerTemp['text']);
  }


  private getListItemType(name: string) {
    let safeListType = "SP.Data." + name[0].toUpperCase() + name.substring(1) + "ListItem";
    safeListType = safeListType.replace(/_/g, "_x005f_");
    safeListType = safeListType.replace(/ /g, "_x0020_");
    return safeListType;
  }

  private createItem() {
    let allFilled=this.isAllFilled();
    var isFairDate=new Date(this.stDate)<=new Date(this.endDate);  
    console.log(isFairDate);
    if (allFilled && isFairDate) {
      
      let requestdata = {};
      requestdata['Title'] = "title";
      requestdata['Description'] = this.desc;
      requestdata['Reason'] = this.reason;
      requestdata['EmpName'] = this.props.currentUserName;
      requestdata['StartDate'] = this.stDate;
      requestdata['EndDate'] = this.endDate;
      requestdata['Status'] = 'Pending';
      requestdata['Manager']=localStorage.getItem("ManagerEmail");
      requestdata['EmpEmail']=this.props.currentUserEmail;
      requestdata['ManagerName']=localStorage.getItem("ManagerName");
      localStorage.removeItem('ManagerEmail');
      localStorage.removeItem('ManagerName');
      console.log(this.life);
      let requestdatastr = JSON.stringify(requestdata);
      requestdatastr = requestdatastr.substring(1, requestdatastr.length - 1);
      console.log(requestdatastr);
      let requestlistItem: string = JSON.stringify({
        '__metadata': { 'type': this.getListItemType(this.props.listName) }
      });
      requestlistItem = requestlistItem.substring(1, requestlistItem.length - 1);
      requestlistItem = '{' + requestlistItem + ',' + requestdatastr + '}';
      console.log(requestlistItem);


      this.props.spHttpClient.post(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: requestlistItem
        }).then((response: SPHttpClientResponse) => {
          //console.log(response.json());
          return response.json();
        }).then((): void => {
          console.log("successfully added...");
          location.reload();
        }, (error: any): void => {
          console.log(error + ": An error occured...");
          alert('An error occured...');
        });       
    }
    else {
      alert("Oops! Looks like mandatory data are not filled or date entered is incorrect... ");
    }

  }
}
