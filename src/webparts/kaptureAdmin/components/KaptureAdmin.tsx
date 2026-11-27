import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IKaptureAdminProps, IKnowledgeFile } from './IKaptureAdminProps';
import styles from './KaptureAdmin.module.scss';

interface IState { files: IKnowledgeFile[]; loading: boolean; uploading: boolean; message: string; error: string; selectedFile?: File; }
export default class KaptureAdmin extends React.Component<IKaptureAdminProps, IState> {
  constructor(props: IKaptureAdminProps) { super(props); this.state = { files: [], loading: true, uploading: false, message: '', error: '' }; }
  public componentDidMount(): void { this.loadFiles(); }
  private api = (path: string): string => `${this.props.siteUrl}/_api${path}`;
  private async loadFiles(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const folder = this.props.repositoryFolder.replace(/'/g, "''");
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(this.api(`/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=TimeLastModified desc`), SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
      if (!response.ok) throw new Error(`Could not access the repository (${response.status}). Confirm the library path and your SharePoint permissions.`);
      const data = await response.json(); this.setState({ files: data.value || [], loading: false });
    } catch (e) { this.setState({ loading: false, error: e instanceof Error ? e.message : 'Unable to load repository documents.' }); }
  }
  private onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => { this.setState({ selectedFile: event.target.files && event.target.files[0] ? event.target.files[0] : undefined }); };
  private async upload(): Promise<void> {
    const file = this.state.selectedFile;
    if (!file) { this.setState({ error: 'Choose a document before uploading.' }); return; }
    this.setState({ uploading: true, error: '', message: '' });
    try {
      const folder = this.props.repositoryFolder.replace(/'/g, "''");
      const safeName = file.name.replace(/'/g, "''");
      const url = this.api(`/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')/Files/add(url='${encodeURIComponent(safeName)}',overwrite=true)`);
      const response = await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' }, body: file });
      if (!response.ok) throw new Error(`Upload failed (${response.status}). Check that you have edit permission and that the file type/size is allowed.`);
      this.setState({ uploading: false, message: 'Document uploaded successfully. It is now in the shared Kapture Knowledge Repository.', selectedFile: undefined });
      const input = document.getElementById('kapture-admin-file') as HTMLInputElement; if (input) input.value = '';
      await this.loadFiles();
    } catch (e) { this.setState({ uploading: false, error: e instanceof Error ? e.message : 'Upload failed.' }); }
  }
  private formatDate(value: string): string { return value ? new Date(value).toLocaleDateString() : '—'; }
  public render(): React.ReactElement<IKaptureAdminProps> {
    return <div className={styles.root}>
      <header className={styles.header}><div><div className={styles.eyebrow}>INDOVANCE · KNOWLEDGE MANAGEMENT</div><h1>Kapture Admin Portal</h1><p>Upload and manage documents in the shared Kapture Knowledge Repository.</p></div><div className={styles.user}>Signed in as <strong>{this.props.displayName}</strong></div></header>
      <div className={styles.notice}><strong>One shared repository</strong><span>Every successful upload is saved directly to the Kapture Knowledge Repository. The Employee Wiki reads from this same library; there is no separate Incoming or Published folder and no additional Publish step.</span></div>
      {this.state.message && <div className={styles.success}>{this.state.message}</div>}{this.state.error && <div className={styles.error}>{this.state.error}</div>}
      <div className={styles.grid}>
        <section className={styles.panel}><div className={styles.panelHead}><div><h2>Upload knowledge</h2><p>Add a document to the shared repository</p></div><span className={styles.step}>UPLOAD</span></div>
          <label className={styles.label}>Select document<input id="kapture-admin-file" type="file" onChange={this.onFileChange} /></label>
          {this.state.selectedFile && <div className={styles.fileSelected}>Selected: {this.state.selectedFile.name}</div>}
          <button className={styles.primary} disabled={this.state.uploading} onClick={() => this.upload()}>{this.state.uploading ? 'Uploading…' : 'Upload document'}</button>
          <p className={styles.helper}>Uploading places the original file directly in the shared SharePoint library. Access is controlled by SharePoint permissions.</p>
        </section>
        <section className={styles.panel}><div className={styles.panelHead}><div><h2>Repository documents</h2><p>Files currently stored in Kapture Knowledge Repository</p></div><button className={styles.refresh} onClick={() => this.loadFiles()}>↻ Refresh</button></div>
          {this.state.loading ? <div className={styles.empty}>Loading documents…</div> : this.state.files.length === 0 ? <div className={styles.empty}>No documents in the repository yet.</div> : <div className={styles.fileList}>{this.state.files.map(file => <div className={styles.file} key={file.ServerRelativeUrl}><div className={styles.fileIcon}>DOC</div><div className={styles.fileInfo}><strong>{file.Name}</strong><span>Modified {this.formatDate(file.TimeLastModified)} · {Math.max(1, Math.round(Number(file.Length || 0) / 1024))} KB</span></div><a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.refresh}>Open</a></div>)}</div>}
          <div className={styles.footerNote}>Documents uploaded here are stored in the same library that the Employee Wiki will use.</div>
        </section>
      </div>
    </div>;
  }
}
