import { SPHttpClient } from '@microsoft/sp-http';
export interface IKaptureAdminProps { siteUrl: string; repositoryFolder: string; spHttpClient: SPHttpClient; displayName: string; }
export interface IKnowledgeFile { Name: string; ServerRelativeUrl: string; TimeLastModified: string; Length: string; Department?: string; SubDepartment?: string; Category?: string; Tags?: string; }
