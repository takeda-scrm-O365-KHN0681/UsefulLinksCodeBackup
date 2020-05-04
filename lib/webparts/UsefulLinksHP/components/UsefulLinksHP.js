var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './UsefulLinksHP.module.scss';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { getTheme, DefaultButton, FontWeights, ContextualMenu, Modal, IconButton } from 'office-ui-fabric-react';
//For Modal dialog
var dragOptions = {
    moveMenuItemText: 'Move',
    closeMenuItemText: 'Close',
    menu: ContextualMenu,
};
var AddIcon = { iconName: 'Add' };
var cancelIcon = { iconName: 'Cancel' };
var EditIcon = { iconName: 'Edit' };
var DeleteIcon = { iconName: 'Delete' };
//End of Modal dialog
var classNames = mergeStyleSets({
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
    iconClass: {
        height: 30,
        width: 20,
        margin: '0 8px',
        color: 'Black'
    }
});
var controlStyles = {
    root: {
        margin: '0 30px 20px 0',
        maxWidth: '700px'
    }
};
//Show modal
var theme = getTheme();
var contentStyles = mergeStyleSets({
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
            borderTop: "4px solid Red",
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
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
var UsefulLinksHP = /** @class */ (function (_super) {
    __extends(UsefulLinksHP, _super);
    function UsefulLinksHP(props, state) {
        var _this = _super.call(this, props) || this;
        _this.showModal = function () {
            _this.setState({ showModal: true });
        };
        _this.closeModal = function () {
            _this.setState({ showModal: false });
        };
        // private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
        //  this.setState({ isCompactMode: checked });
        //}
        //private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
        //  this.setState({ isModalSelection: checked });
        //}
        _this._onChangeText = function (ev, text) {
            _this.setState({
                items: text ? _this.state.items.filter(function (i) { return i.Product.toLowerCase().indexOf(text.toLowerCase()) > -1; }) : _this._allItems,
            });
        };
        _this._onColumnClick = function (ev, column) {
            var _a = _this.state, columns = _a.columns, items = _a.items;
            var newColumns = columns.slice();
            var currColumn = newColumns.filter(function (currCol) { return column.key === currCol.key; })[0];
            newColumns.forEach(function (newCol) {
                if (newCol === currColumn) {
                    currColumn.isSortedDescending = !currColumn.isSortedDescending;
                    currColumn.isSorted = true;
                    _this.setState({
                        announcedMessage: currColumn.name + " is sorted " + (currColumn.isSortedDescending ? 'descending' : 'ascending')
                    });
                }
                else {
                    newCol.isSorted = false;
                    newCol.isSortedDescending = true;
                }
            });
            var newItems = _this._copyAndSort(items, currColumn.fieldName, currColumn.isSortedDescending);
            _this.setState({
                columns: newColumns,
                items: newItems
            });
        };
        _this._allItems = [];
        _this._generateDocuments();
        //this._allItems = _generateDocuments();
        _this._renderEdit = _this._renderEdit.bind(_this);
        _this._renderItemColumn = _this._renderItemColumn.bind(_this);
        _this._DeleteItem = _this._DeleteItem.bind(_this);
        _this._copyAndSort = _this._copyAndSort.bind(_this);
        var columns = [
            {
                key: 'column2',
                name: 'Category',
                fieldName: 'Category',
                minWidth: 150,
                maxWidth: 200,
                headerClassName: styles.TableHeader,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: _this._onColumnClick,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'column3',
                name: 'Name of the Website',
                fieldName: 'Product',
                minWidth: 180,
                maxWidth: 280,
                headerClassName: styles.TableHeader,
                isResizable: true,
                onColumnClick: _this._onColumnClick,
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
                onColumnClick: _this._onColumnClick
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
                onColumnClick: _this._onColumnClick
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
                onColumnClick: _this._onColumnClick
            }
        ];
        _this._selection = new Selection({
            onSelectionChanged: function () {
                _this.setState({
                    selectionDetails: _this._getSelectionDetails()
                });
            }
        });
        _this.state = {
            items: [],
            columns: columns,
            selectionDetails: _this._getSelectionDetails(),
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
        return _this;
    }
    UsefulLinksHP.prototype._renderEdit = function (itemid) {
        console.log("Item value " + itemid);
        //alert(itemid);
        // let currentItem: any = undefined;
        // currentItem.push(this._allItems.filter(e => e.key == itemid.toString()));
        this.setState({ Category: this._allItems[itemid].Category, SiteName: this._allItems[itemid].Product, URL: this._allItems[itemid].Number, IDEdit: this._allItems[itemid].ID });
        this.showModal();
    };
    UsefulLinksHP.prototype.UpdateItem = function (IDEdit) {
        return __awaiter(this, void 0, void 0, function () {
            var weburl, websiteurl, re, Newitempageurl;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        weburl = this.props.SiteUrl;
                        websiteurl = Web(weburl);
                        re = /^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$/;
                        if (!(document.getElementById("Category")["value"] == "" || document.getElementById("URL")["WebsiteName"] == "" || document.getElementById("URL")["value"] == "")) return [3 /*break*/, 1];
                        alert("Error: Please fill out empty fields!");
                        return [2 /*return*/, false];
                    case 1:
                        if (!!re.test(document.getElementById("URL")["value"])) return [3 /*break*/, 2];
                        alert("Error: Input contains invalid URL!");
                        document.getElementById("URL").focus();
                        return [2 /*return*/, false];
                    case 2: return [4 /*yield*/, websiteurl.lists.getByTitle("Useful Links").items.getById(IDEdit).update({
                            Category: document.getElementById('Category')["value"],
                            Title: document.getElementById('WebsiteName')["value"],
                            URL: document.getElementById('URL')["value"]
                        })];
                    case 3:
                        _a.sent();
                        alert("Record of Useful Links has been updated !");
                        this.closeModal();
                        Newitempageurl = "/SitePages/Useful-Links.aspx";
                        window.location.replace(weburl + Newitempageurl);
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    UsefulLinksHP.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var UserPermission;
            return __generator(this, function (_a) {
                UserPermission = this.getUserPermission();
                return [2 /*return*/];
            });
        });
    };
    UsefulLinksHP.prototype.getUserPermission = function () {
        return __awaiter(this, void 0, void 0, function () {
            var userPermission, weburl, websiteurl, userEffectivePermissions;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        userPermission = undefined;
                        weburl = this.props.SiteUrl;
                        websiteurl = Web(weburl);
                        console.log("User permission: " + weburl);
                        return [4 /*yield*/, websiteurl.lists.getByTitle("Useful Links").effectiveBasePermissions.get()];
                    case 1:
                        userEffectivePermissions = _a.sent();
                        this.setState({ hasPermissionAdd: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.AddListItems) });
                        this.setState({ hasPermissionDelete: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.DeleteListItems) });
                        this.setState({ hasPermissionEdit: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.EditListItems) });
                        this.setState({ hasPermissionView: websiteurl.lists.getByTitle("Useful Links").hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems) });
                        return [2 /*return*/];
                }
            });
        });
    };
    UsefulLinksHP.prototype.render = function () {
        var _this = this;
        var _a = this.state, columns = _a.columns, items = _a.items, selectionDetails = _a.selectionDetails, announcedMessage = _a.announcedMessage;
        console.log(columns);
        return (React.createElement("div", { className: styles.itmArticles },
            React.createElement("div", { className: "ms-Grid container" },
                React.createElement("div", { className: styles["ms-Grid-row"] },
                    React.createElement("div", { className: "ms-sm12 ms-md12 ms-lg12 " + styles.monjuTitle },
                        React.createElement("span", { className: styles.customSubLogo }, "UL"),
                        React.createElement("span", null, "Useful Links")))),
            React.createElement(Fabric, null,
                React.createElement("div", { className: styles.controlWrapper },
                    React.createElement("div", { className: styles.inputWrapper },
                        React.createElement(TextField, { label: "Filter by website name:", onChange: this._onChangeText }),
                        React.createElement(Announced, { message: "Number of items after filter applied: " + items.length + "." })),
                    this.state.hasPermissionAdd == true ? React.createElement(DefaultButton, { text: "New Article", iconProps: AddIcon, className: styles.Searchbutton, onClick: function (e) { return _this._CreateDocuments(); } }) : ""),
                React.createElement("div", null),
                (React.createElement(DetailsList, { items: items, 
                    //compact={isCompactMode}
                    columns: columns, selectionMode: SelectionMode.none, getKey: this._getKey, setKey: "none", layoutMode: DetailsListLayoutMode.justified, isHeaderVisible: true, onItemInvoked: this._onItemInvoked, onRenderItemColumn: this._renderItemColumn })),
                React.createElement(Modal, { titleAriaId: "Bookmark Article", isOpen: this.state.showModal, onDismiss: this.closeModal, 
                    //isBlocking={false}
                    //className={contentStyles.container}
                    className: styles.modalWrapper + " " },
                    React.createElement("div", { className: styles.modalWrapperSub },
                        React.createElement("div", { className: contentStyles.body },
                            React.createElement("div", { className: styles.UsefulLinksUIFabric },
                                React.createElement("div", { className: "ms-Grid container" },
                                    React.createElement("div", { className: styles["ms-Grid-row"] },
                                        React.createElement("div", { className: "ms-sm12 ms-md12 ms-lg12 " + styles.pageTitle },
                                            React.createElement("span", null, "Update Useful Links"))),
                                    React.createElement("div", { className: "ITMForm" },
                                        React.createElement("div", { className: "ms-Grid-row " + styles.demoform + " ", dir: "ltr" },
                                            React.createElement("div", { className: styles["section-divider"] }, "Basic Information"),
                                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                                    React.createElement("label", { className: styles.lblNewForm },
                                                        "Category",
                                                        React.createElement(TextField, { type: "text", value: this.state.Category, placeholder: "Enter Category", className: styles.txtboxInput, id: "Category", name: "Category" })))),
                                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                                    React.createElement("label", { className: styles.lblNewForm },
                                                        "Name of the Website",
                                                        React.createElement(TextField, { type: "text", value: this.state.SiteName, placeholder: "Enter Website Name", className: styles.txtboxInput, id: "WebsiteName" })))),
                                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                                    React.createElement("label", { className: styles.lblNewForm },
                                                        "URL",
                                                        React.createElement(TextField, { type: "text", value: this.state.URL, placeholder: "Enter Site URL", className: styles.txtboxInput, id: "URL" })))),
                                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                                    React.createElement("div", { className: styles.btnSection },
                                                        React.createElement("button", { type: "submit", className: styles.btn + " " + styles["btn-publish"], id: "btnsave", onClick: function (e) { _this.UpdateItem(_this.state.IDEdit); } }, "Update"),
                                                        React.createElement("button", { type: "submit", className: styles.btn + " " + styles["btn-cancel"], id: "btnCancel", onClick: this.closeModal }, "Cancel"))))))))))))));
    };
    //items["ID"]
    UsefulLinksHP.prototype.componentDidUpdate = function (previousProps, previousState) {
        // if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
        //  this._selection.setAllSelected(false);
        // }
    };
    UsefulLinksHP.prototype._getKey = function (item, index) {
        return item.key;
    };
    UsefulLinksHP.prototype._onItemInvoked = function (item) {
        alert("Item invoked: " + item.name);
    };
    UsefulLinksHP.prototype._getSelectionDetails = function () {
        var selectionCount = this._selection.getSelectedCount();
        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + this._selection.getSelection()[0].Category;
            default:
                return selectionCount + " items selected";
        }
    };
    UsefulLinksHP.prototype._generateDocuments = function () {
        var _this = this;
        //const items: IDocument[] = [];
        var weburl = this.props.SiteUrl;
        var websiteurl = Web(weburl);
        console.log('Sharepoint Context loaded: ' + weburl);
        var listitem = [];
        websiteurl.lists.getByTitle("Useful Links").items.get().then(function (response) {
            console.log(response);
            for (var i = 0; i < response.length; i++) {
                //console.log(response[i]["Modified"]);
                var str = response[i]["Modified"];
                listitem.push({
                    key: i,
                    Category: response[i]["Category"],
                    Product: response[i]["Title"],
                    Number: response[i]["URL"],
                    PublishedDate: str.substr(0, 10),
                    ID: response[i]["ID"]
                    //Category: response[i]["Category"],
                    // Status: response[i]["Status"]
                });
            }
            _this.setState({ items: listitem });
            _this._allItems = listitem;
            //return items;
        });
    };
    //Row data bound implemented for hyperlink
    UsefulLinksHP.prototype._renderItemColumn = function (item, index, column) {
        var _this = this;
        var fieldContent = item[column.fieldName];
        var DelNumber = item["ID"];
        var EditNumber = item["key"];
        switch (column.key) {
            case 'column9':
                return (React.createElement("span", null,
                    this.state.hasPermissionAdd == true ? React.createElement(IconButton, { ariaLabel: "Edit", iconProps: EditIcon, className: classNames.iconClass, onClick: function (e) { return _this._renderEdit(EditNumber); } }) : "",
                    this.state.hasPermissionAdd == true ? React.createElement(IconButton, { ariaLabel: "Delete", iconProps: DeleteIcon, className: classNames.iconClass, onClick: function () { if (window.confirm('Are you sure you wish to delete this item?')) {
                            _this._DeleteItem(DelNumber);
                        } } }) : ""));
            case 'column4':
                return React.createElement("span", null,
                    React.createElement("a", { href: fieldContent, target: "_blank" }, fieldContent));
            default:
                return React.createElement("span", null, fieldContent);
        }
    };
    UsefulLinksHP.prototype._CreateDocuments = function () {
        var weburl = this.props.SiteUrl;
        var Newitempageurl = "/SitePages/New-Useful-Links.aspx";
        window.location.replace(weburl + Newitempageurl);
    };
    UsefulLinksHP.prototype._DeleteItem = function (DelId) {
        var weburl = this.props.SiteUrl;
        var websiteurl = Web(weburl);
        websiteurl.lists.getByTitle("Useful Links").items.getById(DelId).delete();
        alert("Record with Useful links has been Deleted !");
        //this._generateDocuments();
        var Newitempageurl = "/SitePages/Useful-Links.aspx";
        window.location.replace(weburl + Newitempageurl);
        //this._allItems = this._allItems;
        //this._allItems;
    };
    UsefulLinksHP.prototype._copyAndSort = function (items, columnKey, isSortedDescending) {
        var key = columnKey;
        return items.slice(0).sort(function (a, b) { return ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1); });
    };
    return UsefulLinksHP;
}(React.Component));
export default UsefulLinksHP;
//# sourceMappingURL=UsefulLinksHP.js.map