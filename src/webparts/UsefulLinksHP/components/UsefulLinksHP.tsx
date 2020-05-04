import * as React from 'react';
import styles from './UsefulLinksHP.module.scss';
import { IUsefulLinksHPProps } from './IUsefulLinksHPProps';
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
import { getTheme, DefaultButton, FontWeights, PrimaryButton, ContextualMenu, Stack, Modal, IDragOptions, IconButton, IIconProps, IStackTokens } from 'office-ui-fabric-react';
import { useId, useBoolean } from '@uifabric/react-hooks';

//For Modal dialog
const dragOptions: IDragOptions = {
  moveMenuItemText: 'Move',
  closeMenuItemText: 'Close',
  menu: ContextualMenu,
};
const AddIcon: IIconProps = { iconName: 'Add' };
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const EditIcon: IIconProps = { iconName: 'Edit' };
const DeleteIcon: IIconProps = { iconName: 'Delete' };
//End of Modal dialog

const classNames = mergeStyleSets({
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  headerBackground: {
    background: '#212529',
    color: '#fff;',
    font: '14px',
    height: '30'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  selectionDetails: {
    marginBottom: '20px'
  },
  iconClass:
  {
    height: 30,
    width: 20,
    margin: '0 8px',
    color: 'Black'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '700px'
  }
};
//Show modal
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch'
  },
  containerBook: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    borderTop: '5px solid red'
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      borderTop: `4px solid Red`,
      color: 'Blue',
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px'
    }
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    width: '450px',
    //height: '600px',
    overflowY: 'hidden',
    selectors: {
      p: {
        margin: '14px 0'
      },
      'p:first-child': {
        marginTop: 0
      },
      'p:last-child': {
        marginBottom: 0
      }
    }
  },
  bodybook: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    width: '400px',
    height: '200px',
    overflowY: 'hidden',
    selectors: {
      p: {
        margin: '14px 0'
      },
      'p:first-child': {
        marginTop: 0
      },
      'p:last-child': {
        marginBottom: 0
      }
    }
  }
});

//end show modal
import { Web, PermissionKind } from '@pnp/sp/presets/all';
import { Version } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { _Items, Items } from '@pnp/sp/items/types';
import * as pnp from 'sp-pnp-js';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { Logger } from "@pnp/logging";
import { IUserPermissions } from './IUserPermissions';

export interface IDetailsListDocumentsExampleState {
  columns: IColumn[];
  items: IDocument[];
  selectionDetails: string;
  announcedMessage?: string;
  showModal: boolean;
  hasPermissionAdd: boolean;
  hasPermissionEdit: boolean;
  hasPermissionDelete: boolean;
  hasPermissionView: boolean;
  Category: string;
  SiteName: string;
  URL: string;
  IDEdit: string;
}

export interface IDocument {
  key: string;
  Category: string;
  Product: string;
  Number: string;
  PublishedDate: string;
  ID: string;
  //Category: string;
  //Status: string;
}

export default class UsefulLinksHP extends React.Component<IUsefulLinksHPProps, IDetailsListDocumentsExampleState> {
  private _selection: Selection;
  private _allItems: IDocument[];

  constructor(props: IUsefulLinksHPProps, state: IDetailsListDocumentsExampleState) {
    super(props);

    this._allItems = [];
    this._generateDocuments();
    //this._allItems = _generateDocuments();
    this._renderEdit = this._renderEdit.bind(this);
    this._renderItemColumn = this._renderItemColumn.bind(this);
    this._DeleteItem = this._DeleteItem.bind(this);
    this._copyAndSort = this._copyAndSort.bind(this);

    const columns: IColumn[] = [
      {
        key: 'column2',
        name: 'Category',
        fieldName: 'Category',
        minWidth: 150,
        maxWidth: 200,
        headerClassName: styles.TableHeader, //classNames.headerBackground,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
        //className: styles.TableHeader
      },
      {
        key: 'column3',
        name: 'Name of the Website',
        fieldName: 'Product',
        minWidth: 180,
        maxWidth: 280,
        headerClassName: styles.TableHeader,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column4',
        name: 'URL',
        fieldName: 'Number',
        minWidth: 160,
        maxWidth: 300,
        headerClassName: styles.TableHeader,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick
      },
      {
        key: 'column5',
        name: 'Date Modified',
        fieldName: 'PublishedDate',
        minWidth: 100,
        maxWidth: 120,
        headerClassName: styles.TableHeader,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick
      },
      // {
      //   key: 'column6',
      //   name: 'Category',
      //   fieldName: 'Category',
      //   minWidth: 50,
      //   maxWidth: 70,
      //   headerClassName:classNames.headerBackground,
      //   isResizable: true,
      //   isCollapsible: true,
      //   data: 'string',
      //   onColumnClick: this._onColumnClick
      // },
      // {
      //   key: 'column8',
      //   name: 'Status',
      //   fieldName: 'Status',
      //   minWidth: 50,
      //   maxWidth: 100,
      //   headerClassName:classNames.headerBackground,
      //   isResizable: true,
      //   isCollapsible: true,
      //   data: 'string',
      //   onColumnClick: this._onColumnClick
      // },
      {
        key: 'column9',
        name: 'Actions',
        fieldName: 'Actions',
        minWidth: 10,
        maxWidth: 10,
        headerClassName: styles.TableHeader,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });

    this.state = {
      items: [],
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      //isModalSelection: false,
      //isCompactMode: false,
      announcedMessage: undefined,
      showModal: false,
      hasPermissionAdd: false,
      hasPermissionEdit: false,
      hasPermissionDelete: false,
      hasPermissionView: false,
      Category: undefined,
      //Product: undefined,
      // Number: undefined
      SiteName: undefined,
      URL: undefined,
      IDEdit: undefined
    };
  }

  private showModal = (): void => {
    this.setState({ showModal: true });
  }
  private closeModal = (): void => {
    this.setState({ showModal: false });
  }
  public _renderEdit(itemid: string) {
    console.log("Item value " + itemid);

    //alert(itemid);
    // let currentItem: any = undefined;
    // currentItem.push(this._allItems.filter(e => e.key == itemid.toString()));

    this.setState({ Category: this._allItems[itemid].Category, SiteName: this._allItems[itemid].Product, URL: this._allItems[itemid].Number, IDEdit: this._allItems[itemid].ID });
    this.showModal();
  }

  public async UpdateItem(IDEdit: any) {
    let weburl = this.props.SiteUrl;
    let websiteurl = Web(weburl);
    var re = /^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$/;
    if (document.getElementById("Category")["value"] == "" || document.getElementById("URL")["WebsiteName"] == "" || document.getElementById("URL")["value"] == "") {
      alert("Error: Please fill out empty fields!");
      return false;
    }
    else if (!re.test(document.getElementById("URL")["value"])) {
      alert("Error: Input contains invalid URL!");
      document.getElementById("URL").focus();
      return false;
    }
    else {
      await websiteurl.lists.getByTitle("Useful Links").items.getById(IDEdit).update({
        Category: document.getElementById('Category')["value"],
        Title: document.getElementById('WebsiteName')["value"],
        URL: document.getElementById('URL')["value"]
      });
      alert("Record of Useful Links has been updated !");
      this.closeModal();
      let Newitempageurl = "/SitePages/Useful-Links.aspx";
      window.location.replace(weburl + Newitempageurl);
    }
  }

  public async componentDidMount() {
    const UserPermission: any = this.getUserPermission();
  }

  private async getUserPermission(): Promise<any> {
    let userPermission: IUserPermissions = undefined;

    let weburl = this.props.SiteUrl;
    let websiteurl = Web(weburl);
    console.log("User permission: " + weburl);

    const userEffectivePermissions = await websiteurl.lists.getByTitle("Useful Links").effectiveBasePermissions.get();
    this.setState({ hasPermissionAdd: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.AddListItems) });
    this.setState({ hasPermissionDelete: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.DeleteListItems) });
    this.setState({ hasPermissionEdit: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.EditListItems) });
    this.setState({ hasPermissionView: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems) });
    //userPermission = { hasPermissionAdd: hasPermissionAdd, hasPermissionEdit: hasPermissionEdit, hasPermissionDelete: hasPermissionDelete, hasPermissionView: hasPermissionView };
    //Logger.writeJSON(userPermission);
    //const userEffectivePermissions = await sp.web.effeffectiveBasePermissions.get();
    //console.log("User Permission: " + userPermission);
    //return userPermission;
  }

  public render() {
    const { columns, items, selectionDetails, announcedMessage } = this.state;
    console.log(columns);
    return (

      <div className={styles.itmArticles}>
        <div className="ms-Grid container">
          <div className={styles["ms-Grid-row"]}>
            <div className={`ms-sm12 ms-md12 ms-lg12 ${styles.monjuTitle}`}>
              <span className={styles.customSubLogo}>UL</span>
              <span>Useful Links</span>
            </div>
          </div>
        </div>

        <Fabric>
          <div className={styles.controlWrapper}>
            <div className={styles.inputWrapper}>
              <TextField label="Filter by website name:" onChange={this._onChangeText} />
              <Announced message={`Number of items after filter applied: ${items.length}.`} />
            </div>
            {/*<DefaultButton text="New Article" className={styles.Searchbutton} onClick={() => { console.log('Button clicked'); }} />*/}
            {/*{this.state.hasPermissionAdd == true ? <DefaultButton text="New Article" className={styles.Searchbutton} onClick={() => { console.log('Button clicked'); }} /> : ""}*/}
            {this.state.hasPermissionAdd == true ? <DefaultButton text="New Article" iconProps={AddIcon} className={styles.Searchbutton} onClick={e => this._CreateDocuments()} /> : ""}
          </div>
          <div></div>
          {(
            <DetailsList
              items={items}
              //compact={isCompactMode}
              columns={columns}
              selectionMode={SelectionMode.none}
              getKey={this._getKey}
              setKey="none"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              onItemInvoked={this._onItemInvoked}
              onRenderItemColumn={this._renderItemColumn}
            />
          )}
          <Modal
            titleAriaId={"Bookmark Article"}
            isOpen={this.state.showModal}
            onDismiss={this.closeModal}
            //isBlocking={false}
            //className={contentStyles.container}
            className={`${styles.modalWrapper} `}

          >

            <div className={styles.modalWrapperSub}>
              <div className={contentStyles.body}>
                <div className={styles.UsefulLinksUIFabric}>
                  <div className="ms-Grid container">
                    <div className={styles["ms-Grid-row"]}>
                      {/* <div className={`ms-sm12 ms-md12 ms-lg12 ${styles.monjuTitle}`}>
                        <span className={styles.ModalTitle}>UL</span>
                        <span>Update Useful Links</span>

                      </div> */}
                      <div className={`ms-sm12 ms-md12 ms-lg12 ${styles.pageTitle}`}>
                        {/* <span className={styles.customSubLogo}>UL</span> */}
                        <span>Update Useful Links</span>
                      </div>
                    </div>
                    <div className="ITMForm" >

                      {/* </div><div className="ms-Grid-row demoform shadow-sm" dir="ltr"> */}
                      <div className={`ms-Grid-row ${styles.demoform} `} dir="ltr">
                        <div className={styles["section-divider"]}>
                          Basic Information
                </div>
                        <div className={styles["ms-Grid-row"]}>
                          {/*<div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"> */}
                          <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                            <label className={styles.lblNewForm}>Category
					 <TextField type="text" value={this.state.Category} placeholder="Enter Category" className={styles.txtboxInput} id="Category" name="Category" />
                              {/* <TextField floatingLabelText={this.props.heading} type={this.props.inputType} value={this.props.value}  onChange={this.props._change} /> */}

                            </label>
                          </div>
                        </div>

                        <div className={styles["ms-Grid-row"]}>
                          {/*<div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"> */}
                          <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                            <label className={styles.lblNewForm}>Name of the Website
					 <TextField type="text" value={this.state.SiteName} placeholder="Enter Website Name" className={styles.txtboxInput} id="WebsiteName" />
                            </label>
                          </div>
                        </div>

                        <div className={styles["ms-Grid-row"]}>
                          <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                            <label className={styles.lblNewForm}>URL
					 <TextField type="text" value={this.state.URL} placeholder="Enter Site URL" className={styles.txtboxInput} id="URL" />
                            </label>
                          </div>
                        </div>

                        {/* <div className={styles["ms-Grid-row"]}>
                          <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                            <label className={styles.lblNewForm}>URL
					 <TextField type="text" value={this.state.IDEdit} placeholder="" className={styles.txtboxInput} id="ID" />
                              </label>
                          </div>
                        </div> */}

                        <div className={styles["ms-Grid-row"]}>
                          <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                            <div className={styles.btnSection}>

                              <button type="submit" className={`${styles.btn} ${styles["btn-publish"]}`} id="btnsave" onClick={(e) => { this.UpdateItem(this.state.IDEdit); }}>Update</button>
                              <button type="submit" className={`${styles.btn} ${styles["btn-cancel"]}`} id="btnCancel" onClick={this.closeModal}>Cancel</button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </Modal>
        </Fabric>
      </div>
    );
  }
  //items["ID"]
  public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
    // if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
    //  this._selection.setAllSelected(false);
    // }
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  // private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
  //  this.setState({ isCompactMode: checked });
  //}

  //private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
  //  this.setState({ isModalSelection: checked });
  //}

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this.state.items.filter(i => i.Product.toLowerCase().indexOf(text.toLowerCase()) > -1) : this._allItems,
    });
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).Category;
      default:
        return `${selectionCount} items selected`;
    }
  }
  public _generateDocuments() {
    //const items: IDocument[] = [];
    let weburl = this.props.SiteUrl;
    let websiteurl = Web(weburl);
    console.log('Sharepoint Context loaded: ' + weburl);
    let listitem: any[] = [];

    websiteurl.lists.getByTitle("Useful Links").items.get().then((response) => {
      console.log(response);
      for (let i = 0; i < response.length; i++) {
        //console.log(response[i]["Modified"]);
        var str = response[i]["Modified"];
        listitem.push({
          key: i,
          Category: response[i]["Category"],
          Product: response[i]["Title"],
          Number: response[i]["URL"], //["Url"],
          PublishedDate: str.substr(0, 10),
          ID: response[i]["ID"]
          //Category: response[i]["Category"],
          // Status: response[i]["Status"]
        });
      }
      this.setState({ items: listitem });
      this._allItems = listitem;
      //return items;
    });
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'}`
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
  //Row data bound implemented for hyperlink
  private _renderItemColumn(item: IColumn, index: number, column: IColumn) {
    const fieldContent = item[column.fieldName as keyof IColumn] as string;
    const DelNumber = item["ID"];
    const EditNumber = item["key"];

    switch (column.key) {
      case 'column9':
        return (
          <span>
            {this.state.hasPermissionAdd == true ? <IconButton ariaLabel="Edit" iconProps={EditIcon} className={classNames.iconClass} onClick={e => this._renderEdit(EditNumber)} /> : ""}
            {this.state.hasPermissionAdd == true ? <IconButton ariaLabel="Delete" iconProps={DeleteIcon} className={classNames.iconClass} onClick={() => { if (window.confirm('Are you sure you wish to delete this item?')) { this._DeleteItem(DelNumber); } }} /> : ""}

            {/* {this.state.hasPermissionAdd == true ? "" : <IconButton ariaLabel="Edit" iconProps={EditIcon} className={classNames.iconClass} onClick={e => this._renderEdit({ key: e })} />} */}
            {/*<Icon iconName="Delete" className={classNames.iconClass} onClick={() => { if (window.confirm('Are you sure you wish to delete this item?')) { DeleteItem(DelNumber); } }} />*/}
            {/*<Icon iconName="View" className={classNames.iconClass} />*/}
            {/*<Icon iconName="Edit" className={classNames.iconClass} id="btnEdit" onClick={e => this._renderEdit({ key: e })} /> */}
          </span>
        );

      case 'column4':
        return <span><a href={fieldContent} target="_blank">{fieldContent}</a></span>;
      default:
        return <span>{fieldContent}</span>;
    }
  }
  private _CreateDocuments() {
    let weburl = this.props.SiteUrl;
    let Newitempageurl = "/SitePages/New-Useful-Links.aspx";
    window.location.replace(weburl + Newitempageurl);
  }

  private _DeleteItem(DelId: number) {
    let weburl = this.props.SiteUrl;
    let websiteurl = Web(weburl);

    websiteurl.lists.getByTitle("Useful Links").items.getById(DelId).delete();
    alert("Record with Useful links has been Deleted !");
    //this._generateDocuments();
    let Newitempageurl = "/SitePages/Useful-Links.aspx";
    window.location.replace(weburl + Newitempageurl);
    //this._allItems = this._allItems;
    //this._allItems;
  }

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }
}









