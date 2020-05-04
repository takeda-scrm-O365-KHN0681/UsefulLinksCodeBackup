import * as React from 'react';
import { IUsefulLinksHPProps } from './IUsefulLinksHPProps';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
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
}
export default class UsefulLinksHP extends React.Component<IUsefulLinksHPProps, IDetailsListDocumentsExampleState> {
    private _selection;
    private _allItems;
    constructor(props: IUsefulLinksHPProps, state: IDetailsListDocumentsExampleState);
    private showModal;
    private closeModal;
    _renderEdit(itemid: string): void;
    UpdateItem(IDEdit: any): Promise<boolean>;
    componentDidMount(): Promise<void>;
    private getUserPermission;
    render(): JSX.Element;
    componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState): void;
    private _getKey;
    private _onChangeText;
    private _onItemInvoked;
    private _getSelectionDetails;
    _generateDocuments(): void;
    private _onColumnClick;
    private _renderItemColumn;
    private _CreateDocuments;
    private _DeleteItem;
    private _copyAndSort;
}
//# sourceMappingURL=UsefulLinksHP.d.ts.map