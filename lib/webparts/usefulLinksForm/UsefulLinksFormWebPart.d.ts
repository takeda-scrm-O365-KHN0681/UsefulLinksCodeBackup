import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IUsefulLinksFormWebPartProps {
    description: string;
    SiteUrl: string;
    context: WebPartContext;
}
export default class UsefulLinksFormWebPart extends BaseClientSideWebPart<IUsefulLinksFormWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=UsefulLinksFormWebPart.d.ts.map