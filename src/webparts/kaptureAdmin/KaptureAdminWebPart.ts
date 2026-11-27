import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import KaptureAdmin from './components/KaptureAdmin';
import { IKaptureAdminProps } from './components/IKaptureAdminProps';

export interface IKaptureAdminWebPartProps { repositoryFolder: string; }
export default class KaptureAdminWebPart extends BaseClientSideWebPart<IKaptureAdminWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IKaptureAdminProps> = React.createElement(KaptureAdmin, {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      repositoryFolder: this.properties.repositoryFolder || '/sites/Kapture/Shared Documents/Kapture Knowledge Repository',
      spHttpClient: this.context.spHttpClient,
      displayName: this.context.pageContext.user.displayName || 'Kapture administrator'
    });
    ReactDom.render(element, this.domElement);
  }
  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration() { return { pages: [{ header: { description: 'Configure the single shared Kapture document library' }, groups: [{ groupName: 'SharePoint repository', groupFields: [
    PropertyPaneTextField('repositoryFolder', { label: 'Document library server-relative URL' })
  ] }] }] }; }
}
