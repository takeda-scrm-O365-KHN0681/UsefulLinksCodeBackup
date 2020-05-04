import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IUsefulLinksHPProps } from './components/IUsefulLinksHPProps';
export interface IUsefulLinksHPWebPartProps {
    description: string;
    SiteUrl: string;
}
export default class UsefulLinksHPWebPart extends BaseClientSideWebPart<IUsefulLinksHPProps> {
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=UsefulLinksHPWebPart.d.ts.map