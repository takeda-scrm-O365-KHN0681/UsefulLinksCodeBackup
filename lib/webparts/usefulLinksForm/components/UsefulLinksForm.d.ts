import * as React from 'react';
import { IUsefulLinksFormProps } from './IUsefulLinksFormProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface IUsefulLinksFormState {
    Category: string;
    WebsiteName: string;
    URL: String;
}
export default class UsefulLinksForm extends React.Component<IUsefulLinksFormProps, IUsefulLinksFormState> {
    constructor(props: IUsefulLinksFormProps, state: IUsefulLinksFormState);
    render(): React.ReactElement<IUsefulLinksFormProps>;
    _cancelForm(): void;
    AddItem(): Promise<void>;
    checkForm(): boolean;
}
//# sourceMappingURL=UsefulLinksForm.d.ts.map