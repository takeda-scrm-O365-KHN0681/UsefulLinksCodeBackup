import * as React from 'react';
import styles from './UsefulLinksForm.module.scss';
import { IUsefulLinksFormProps } from './IUsefulLinksFormProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton, Stack, IStackTokens } from 'office-ui-fabric-react';

import * as pnp from 'sp-pnp-js';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
//import { Web } from '@pnp/sp/webs';
import { Web } from '@pnp/sp/presets/all';
import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IUsefulLinksFormState {
  Category: string;
  WebsiteName: string;
  URL: String;
}

export default class UsefulLinksForm extends React.Component<IUsefulLinksFormProps, IUsefulLinksFormState> {
  constructor(props: IUsefulLinksFormProps, state: IUsefulLinksFormState) {
    super(props);
    this._cancelForm = this._cancelForm.bind(this);
  }
 
  public render(): React.ReactElement<IUsefulLinksFormProps> {

    return (
      <form action="form-handler" onSubmit={(e) => { this.checkForm(); e.preventDefault(); }} >
        <div className={styles.usefulLinksForm}>
          <div className="ms-Grid container">
            <div className={styles["ms-Grid-row"]}>
              <div className={`ms-sm12 ms-md12 ms-lg12 ${styles.monjuTitle}`}>
                <span className={styles.customSubLogo}>UL</span>
                <span>New Useful Links</span>

              </div>
            </div>
            {/* <div className={styles["ms-Grid-row"]}>
              <div className={`ms-sm12 ms-md12 ms-lg12 ${styles.backButtonWrapper}`}>
                <span className={styles.backButton}></span>
                <span className={styles.backButtonText}>BACK</span>

              </div>

            </div> */}

            <div className="ITMForm" >

              {/* </div><div className="ms-Grid-row demoform shadow-sm" dir="ltr"> */}
              <div className={`ms-Grid-row ${styles.demoform} ${styles["shadow-sm"]}`} dir="ltr">
                <div className={styles["section-divider"]}>
                  Basic Information
                </div>
                <div className={styles["ms-Grid-row"]}>
                  {/*<div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"> */}
                  <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                    <label className={styles.lblNewForm}>Category
					 <input type="text" placeholder="Enter Category" className={styles.txtboxInput} id="Category" required>
                      </input></label>
                  </div>
                </div>

                <div className={styles["ms-Grid-row"]}>
                  {/*<div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"> */}
                  <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                    <label className={styles.lblNewForm}>Name of the Website
					 <input type="text" placeholder="Enter Website Name" className={styles.txtboxInput} id="WebsiteName" required>
                      </input></label>
                  </div>
                </div>

                <div className={styles["ms-Grid-row"]}>
                  <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                    <label className={styles.lblNewForm}>URL
					 <input type="text" placeholder="Enter Site URL" className={styles.txtboxInput} id="URL" required>
                      </input></label>
                  </div>
                </div>

                <div className={styles["ms-Grid-row"]}>
                  <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                    <div className={styles.btnSection}>
                      <button type="submit" className={`${styles.btn} ${styles["btn-publish"]}`} id="btnsave" onClick={this.checkForm} >Save</button>
                      {/* <button type="submit" className={`${styles.btn} ${styles["btn-publish"]}`} id="btnsave" onClick={this.checkForm} >Update</button> */}
                      <button type="reset" className={`${styles.btn} ${styles["btn-cancel"]}`} id="btnCancel" onClick={this._cancelForm}>Cancel</button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </form >
    );
  }
  
  public _cancelForm() {
    let weburl = this.props.SiteUrl;
    let Newitempageurl = "/SitePages/Useful-Links.aspx";
    window.location.replace(weburl + Newitempageurl);
  }

  public async AddItem() {
    let weburl = this.props.SiteUrl;
    let websiteurl = Web(weburl);
    
    await websiteurl.lists.getByTitle("Useful Links").items.add({
      Category: document.getElementById('Category')["value"],
      Title: document.getElementById('WebsiteName')["value"],
      URL: document.getElementById('URL')["value"]
    });
    alert("Record with Useful Links Added !");
    let Newitempageurl = "/SitePages/Useful-Links.aspx";
    window.location.replace(weburl + Newitempageurl);
    //let Newitempageurl = "https://mytakeda.sharepoint.com.rproxy.goskope.com/sites/HomeNaviDev/_layouts/15/workbench.aspx";
    //window.location.replace(Newitempageurl);
  }

  // public UpdateItem(): void {
  //   var id = 85; //document.getElementById('EmployeeId')["value"];
  //   pnp.sp.web.lists.getByTitle("Useful Links").items.getById(id).update({
  //     Category: document.getElementById('Category')["value"],
  //     Title: document.getElementById('WebsiteName')["value"],
  //     URL: document.getElementById('URL')["value"]
  //   });
  //   alert("Record with id 85 has been updated !");
  // }

  public checkForm() {
    //var re = /^[\w ]+$/;
    var re = /^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$/;

    if (document.getElementById("URL")["value"] != "" && !re.test(document.getElementById("URL")["value"])) {
      alert("Error: Input contains invalid URL!");
      document.getElementById("URL").focus();
      return false;
    }
    else {
      this.AddItem();
      return true;
    }
  }
 
}
  