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
import styles from './UsefulLinksForm.module.scss';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
//import { Web } from '@pnp/sp/webs';
import { Web } from '@pnp/sp/presets/all';
var UsefulLinksForm = /** @class */ (function (_super) {
    __extends(UsefulLinksForm, _super);
    function UsefulLinksForm(props, state) {
        var _this = _super.call(this, props) || this;
        _this._cancelForm = _this._cancelForm.bind(_this);
        return _this;
    }
    UsefulLinksForm.prototype.render = function () {
        var _this = this;
        return (React.createElement("form", { action: "form-handler", onSubmit: function (e) { _this.checkForm(); e.preventDefault(); } },
            React.createElement("div", { className: styles.usefulLinksForm },
                React.createElement("div", { className: "ms-Grid container" },
                    React.createElement("div", { className: styles["ms-Grid-row"] },
                        React.createElement("div", { className: "ms-sm12 ms-md12 ms-lg12 " + styles.monjuTitle },
                            React.createElement("span", { className: styles.customSubLogo }, "UL"),
                            React.createElement("span", null, "New Useful Links"))),
                    React.createElement("div", { className: "ITMForm" },
                        React.createElement("div", { className: "ms-Grid-row " + styles.demoform + " " + styles["shadow-sm"], dir: "ltr" },
                            React.createElement("div", { className: styles["section-divider"] }, "Basic Information"),
                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                    React.createElement("label", { className: styles.lblNewForm },
                                        "Category",
                                        React.createElement("input", { type: "text", placeholder: "Enter Category", className: styles.txtboxInput, id: "Category", required: true })))),
                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                    React.createElement("label", { className: styles.lblNewForm },
                                        "Name of the Website",
                                        React.createElement("input", { type: "text", placeholder: "Enter Website Name", className: styles.txtboxInput, id: "WebsiteName", required: true })))),
                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                    React.createElement("label", { className: styles.lblNewForm },
                                        "URL",
                                        React.createElement("input", { type: "text", placeholder: "Enter Site URL", className: styles.txtboxInput, id: "URL", required: true })))),
                            React.createElement("div", { className: styles["ms-Grid-row"] },
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md12 ms-lg12" },
                                    React.createElement("div", { className: styles.btnSection },
                                        React.createElement("button", { type: "submit", className: styles.btn + " " + styles["btn-publish"], id: "btnsave", onClick: this.checkForm }, "Save"),
                                        React.createElement("button", { type: "reset", className: styles.btn + " " + styles["btn-cancel"], id: "btnCancel", onClick: this._cancelForm }, "Cancel"))))))))));
    };
    UsefulLinksForm.prototype._cancelForm = function () {
        var weburl = this.props.SiteUrl;
        var Newitempageurl = "/SitePages/Useful-Links.aspx";
        window.location.replace(weburl + Newitempageurl);
    };
    UsefulLinksForm.prototype.AddItem = function () {
        return __awaiter(this, void 0, void 0, function () {
            var weburl, websiteurl, Newitempageurl;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        weburl = this.props.SiteUrl;
                        websiteurl = Web(weburl);
                        return [4 /*yield*/, websiteurl.lists.getByTitle("Useful Links").items.add({
                                Category: document.getElementById('Category')["value"],
                                Title: document.getElementById('WebsiteName')["value"],
                                URL: document.getElementById('URL')["value"]
                            })];
                    case 1:
                        _a.sent();
                        alert("Record with Useful Links Added !");
                        Newitempageurl = "/SitePages/Useful-Links.aspx";
                        window.location.replace(weburl + Newitempageurl);
                        return [2 /*return*/];
                }
            });
        });
    };
    // public UpdateItem(): void {
    //   var id = 85; //document.getElementById('EmployeeId')["value"];
    //   pnp.sp.web.lists.getByTitle("Useful Links").items.getById(id).update({
    //     Category: document.getElementById('Category')["value"],
    //     Title: document.getElementById('WebsiteName')["value"],
    //     URL: document.getElementById('URL')["value"]
    //   });
    //   alert("Record with id 85 has been updated !");
    // }
    UsefulLinksForm.prototype.checkForm = function () {
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
    };
    return UsefulLinksForm;
}(React.Component));
export default UsefulLinksForm;
//# sourceMappingURL=UsefulLinksForm.js.map