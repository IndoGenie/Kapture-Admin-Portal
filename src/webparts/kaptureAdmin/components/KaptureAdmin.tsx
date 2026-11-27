import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IKaptureAdminProps, IKnowledgeFile } from './IKaptureAdminProps';
import styles from './KaptureAdmin.module.scss';

const DEPARTMENTS: string[] = ['HR', 'Finance', 'Marketing', 'Sales', 'Training and Development', 'IT', 'BES', 'AI', 'Technology', 'Operations'];
const OPERATIONS: string[] = ['Land Survey', 'Land Development', 'Structural', 'Visual Engineering', 'Mechanical', 'Building Services', 'Construction Engineering', 'Water and Drainage', 'Energy and Utilities', 'Heavy Civil and Infra'];
const CATEGORIES: string[] = ['SOP', 'Policy', 'Standards', 'Training Material', 'Video & Recorded Sessions', 'Work Instructions', 'Guidelines', 'Templates & Forms', 'Project Reference', 'Technical Documentation', 'Other'];

interface IState {
  files: IKnowledgeFile[]; loading: boolean; uploading: boolean; message: string; error: string;
  selectedFile?: File; department: string; subDepartment: string; category: string; description: string; tags: string; search: string;
}

export default class KaptureAdmin extends React.Component<IKaptureAdminProps, IState> {
  constructor(props: IKaptureAdminProps) {
    super(props);
    this.state = { files: [], loading: true, uploading: false, message: '', error: '', department: '', subDepartment: '', category: '', description: '', tags: '', search: '' };
  }
  public componentDidMount(): void { this.loadFiles(); }
  private api = (path: string): string => `${this.props.siteUrl}/_api${path}`;
  private escapeOData = (value: string): string => value.replace(/'/g, "''");
  private async request(url: string, method: string, body?: string | Blob): Promise<SPHttpClientResponse> {
    const headers: { [key: string]: string } = { Accept: 'application/json;odata=nometadata' };
    if (method !== 'GET') headers['Content-Type'] = body instanceof Blob ? 'application/octet-stream' : 'application/json;odata=nometadata';
    return this.props.spHttpClient.fetch(url, SPHttpClient.configurations.v1, { method, headers, body: body as any });
  }
  private async ensureMetadataFields(): Promise<void> {
    const fields = [
      { name: 'KaptureDepartment', title: 'Kapture Department' },
      { name: 'KaptureSubDepartment', title: 'Kapture Sub-department' },
      { name: 'KaptureCategory', title: 'Kapture Category' },
      { name: 'KaptureDescription', title: 'Kapture Description' },
      { name: 'KaptureTags', title: 'Kapture Tags' }
    ];
    for (const field of fields) {
      const check = await this.request(this.api(`/web/lists/getbytitle('Documents')/fields/getbyinternalnameortitle('${field.name}')`), 'GET');
      if (check.ok) continue;
      const xml = `<Field Type='Note' DisplayName='${field.title}' Name='${field.name}' StaticName='${field.name}' Group='Kapture Knowledge' />`;
      const add = await this.request(this.api("/web/lists/getbytitle('Documents')/fields/addFieldAsXml"), 'POST', JSON.stringify({ parameters: { __metadata: { type: 'SP.FieldCreationInformation' }, SchemaXml: xml, Options: 0 } }));
      if (!add.ok) throw new Error(`Could not create the '${field.title}' metadata column in the Documents library. Ask the SharePoint administrator to create it or grant the required permission. (${add.status})`);
    }
  }
  private async ensureFolder(path: string): Promise<void> {
    const check = await this.request(this.api(`/web/GetFolderByServerRelativeUrl('${this.escapeOData(path)}')`), 'GET');
    if (check.ok) return;
    const add = await this.request(this.api(`/web/folders/add('${this.escapeOData(path)}')`), 'POST', JSON.stringify({}));
    if (!add.ok) throw new Error(`Could not create repository folder '${path}'. Check SharePoint permissions. (${add.status})`);
  }
  private async ensureDestination(): Promise<string> {
    const root = this.props.repositoryFolder.replace(/\/$/, '');
    await this.ensureFolder(root);
    let destination = root;
    if (this.state.department === 'Operations') {
      destination += '/Operations'; await this.ensureFolder(destination);
      destination += `/${this.state.subDepartment}`; await this.ensureFolder(destination);
    } else {
      destination += `/${this.state.department}`; await this.ensureFolder(destination);
    }
    return destination;
  }
  private async loadFiles(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const root = this.props.repositoryFolder.replace(/'/g, "''");
      const response = await this.request(this.api(`/web/lists/getbytitle('Documents')/items?$select=File/Name,File/ServerRelativeUrl,File/Length,Modified&$expand=File&$filter=FSObjType eq 0&$orderby=Modified desc&$top=500`), 'GET');
      if (!response.ok) throw new Error(`Could not access repository documents (${response.status}). Confirm the Documents library and your SharePoint permissions.`);
      const data = await response.json();
      const files: IKnowledgeFile[] = (data.value || []).map((item: any) => ({ Name: item.File.Name, ServerRelativeUrl: item.File.ServerRelativeUrl, Length: String(item.File.Length || 0), TimeLastModified: item.Modified })).filter((file: IKnowledgeFile) => file.ServerRelativeUrl && file.ServerRelativeUrl.indexOf(root) === 0);
      this.setState({ files, loading: false });
    } catch (e) { this.setState({ loading: false, error: e instanceof Error ? e.message : 'Unable to load repository documents.' }); }
  }
  private onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => { this.setState({ selectedFile: event.target.files && event.target.files[0] ? event.target.files[0] : undefined, message: '', error: '' }); };
  private async upload(): Promise<void> {
    const file = this.state.selectedFile;
    if (!file) { this.setState({ error: 'Choose a document or video before uploading.' }); return; }
    if (!this.state.department || !this.state.category) { this.setState({ error: 'Select a department and document category.' }); return; }
    if (this.state.department === 'Operations' && !this.state.subDepartment) { this.setState({ error: 'Select an Operations sub-department.' }); return; }
    this.setState({ uploading: true, error: '', message: '' });
    try {
      await this.ensureMetadataFields();
      const destination = await this.ensureDestination();
      const safeName = this.escapeOData(file.name);
      const uploadUrl = this.api(`/web/GetFolderByServerRelativeUrl('${this.escapeOData(destination)}')/Files/add(url='${safeName}',overwrite=true)`);
      const uploadResponse = await this.request(uploadUrl, 'POST', file);
      if (!uploadResponse.ok) throw new Error(`Upload failed (${uploadResponse.status}). Check file size, file type and SharePoint edit permissions.`);
      const uploaded = await uploadResponse.json();
      const fileUrl = uploaded.ServerRelativeUrl || `${destination}/${file.name}`;
      const metadata = [
        { FieldName: 'KaptureDepartment', FieldValue: this.state.department },
        { FieldName: 'KaptureSubDepartment', FieldValue: this.state.subDepartment },
        { FieldName: 'KaptureCategory', FieldValue: this.state.category },
        { FieldName: 'KaptureDescription', FieldValue: this.state.description },
        { FieldName: 'KaptureTags', FieldValue: this.state.tags }
      ];
      const metadataResponse = await this.request(this.api(`/web/GetFileByServerRelativeUrl('${this.escapeOData(fileUrl)}')/ListItemAllFields/ValidateUpdateListItem`), 'POST', JSON.stringify({ formValues: metadata, bNewDocumentUpdate: false }));
      if (!metadataResponse.ok) throw new Error('The file was uploaded, but its metadata could not be saved. Please ask the SharePoint administrator to check the Kapture metadata columns.');
      this.setState({ uploading: false, message: `Successfully uploaded ${file.name} to ${this.state.department}${this.state.subDepartment ? ` / ${this.state.subDepartment}` : ''}.`, selectedFile: undefined, description: '', tags: '' });
      const input = document.getElementById('kapture-admin-file') as HTMLInputElement; if (input) input.value = '';
      await this.loadFiles();
    } catch (e) { this.setState({ uploading: false, error: e instanceof Error ? e.message : 'Upload failed.' }); }
  }
  private formatDate(value: string): string { return value ? new Date(value).toLocaleDateString() : '—'; }
  private get visibleFiles(): IKnowledgeFile[] { const term = this.state.search.toLowerCase(); return this.state.files.filter(f => f.Name.toLowerCase().indexOf(term) >= 0); }
  public render(): React.ReactElement<IKaptureAdminProps> {
    const isVideo = !!this.state.selectedFile && this.state.selectedFile.type.indexOf('video/') === 0;
    return <div className={styles.root}>
      <div className={styles.topbar}><div className={styles.brandMark}>K</div><div className={styles.brandText}><strong>KAPTURE</strong><span>KNOWLEDGE MANAGEMENT</span></div><div className={styles.topSpacer} /><div className={styles.profile}><span className={styles.avatar}>{(this.props.displayName || 'A').charAt(0).toUpperCase()}</span><span><small>ADMINISTRATOR</small><strong>{this.props.displayName}</strong></span></div></div>
      <section className={styles.hero}><div className={styles.heroCopy}><div className={styles.eyebrow}>INDOVANCE · KNOWLEDGE OPERATIONS</div><h1>Knowledge Command Centre</h1><p>One place to curate, organise and share the knowledge that moves Indovance forward.</p><div className={styles.heroMeta}><span><i className={styles.liveDot} /> Connected to Kapture SharePoint</span><span>·</span><span>Shared knowledge repository</span></div></div><div className={styles.heroArt}><div className={styles.artCircle}><span>k</span></div><div className={styles.artTile}><b>KNOWLEDGE</b><small>CAPTURE · ORGANISE · SHARE</small></div></div></section>
      <div className={styles.sectionTitle}><div><span className={styles.kicker}>WORKSPACE</span><h2>Manage your knowledge</h2><p>Upload new resources or find what is already available.</p></div><button className={styles.refresh} onClick={() => this.loadFiles()}>↻ Refresh repository</button></div>
      {this.state.message && <div className={styles.success}><b>✓</b>{this.state.message}</div>}{this.state.error && <div className={styles.error}><b>!</b>{this.state.error}</div>}
      <div className={styles.layout}>
        <section className={styles.uploadCard}><div className={styles.cardHead}><div className={styles.iconBadge}><span>↑</span></div><div><h3>Add to the knowledge base</h3><p>Share a useful resource with your colleagues.</p></div></div>
          <label className={styles.dropzone}><input id="kapture-admin-file" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.mp4,.mov,.wmv,.avi,.m4v,.webm" onChange={this.onFileChange} /><span className={styles.uploadGlyph}>⇧</span><strong>{this.state.selectedFile ? this.state.selectedFile.name : 'Choose a file to upload'}</strong><small>{this.state.selectedFile ? `${(this.state.selectedFile.size / (1024 * 1024)).toFixed(2)} MB · ${isVideo ? 'Video' : 'Document'}` : 'Drag-and-drop is coming soon · Browse your device'}</small><em>Browse files</em></label>
          <div className={styles.formGrid}><label className={styles.field}>Department<select value={this.state.department} onChange={e => this.setState({ department: e.target.value, subDepartment: '' })}><option value="">Select department</option>{DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>
            {this.state.department === 'Operations' && <label className={styles.field}>Operations sub-department<select value={this.state.subDepartment} onChange={e => this.setState({ subDepartment: e.target.value })}><option value="">Select sub-department</option>{OPERATIONS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>}
            <label className={styles.field}>Document category<select value={this.state.category} onChange={e => this.setState({ category: e.target.value })}><option value="">Select category</option>{CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}</select></label>
            <label className={`${styles.field} ${styles.full}`}>Description <textarea rows={2} value={this.state.description} onChange={e => this.setState({ description: e.target.value })} placeholder="What is this resource about? (optional)" /></label>
            <label className={`${styles.field} ${styles.full}`}>Tags <input value={this.state.tags} onChange={e => this.setState({ tags: e.target.value })} placeholder="e.g. design, onboarding, quality (comma-separated)" /></label></div>
          <button className={styles.primary} disabled={this.state.uploading} onClick={() => this.upload()}>{this.state.uploading ? <><span className={styles.spinner} /> Uploading resource…</> : 'Upload to knowledge repository  →'}</button>
          <div className={styles.secureNote}><span>▣</span> Files are stored in SharePoint. Access remains governed by your organisation's permissions.</div>
        </section>
        <section className={styles.libraryCard}><div className={styles.libraryHead}><div><span className={styles.kicker}>YOUR CONTENT</span><h3>Repository documents</h3><p>Latest files in the repository root.</p></div><div className={styles.count}>{this.state.files.length}<small>FILES</small></div></div>
          <div className={styles.searchBox}><span>⌕</span><input value={this.state.search} onChange={e => this.setState({ search: e.target.value })} placeholder="Search documents by name…" /></div>
          {this.state.loading ? <div className={styles.empty}>Loading repository…</div> : this.visibleFiles.length === 0 ? <div className={styles.empty}><span>▤</span><strong>{this.state.files.length ? 'No matching documents' : 'Your knowledge library is ready'}</strong><small>{this.state.files.length ? 'Try another search term.' : 'Your uploaded resources will appear here.'}</small></div> : <div className={styles.fileList}>{this.visibleFiles.map(file => <div className={styles.file} key={file.ServerRelativeUrl}><div className={styles.fileIcon}>{file.Name.toLowerCase().match(/\.(mp4|mov|wmv|avi|m4v|webm)$/) ? '▶' : 'DOC'}</div><div className={styles.fileInfo}><strong>{file.Name}</strong><span>Modified {this.formatDate(file.TimeLastModified)} · {Math.max(1, Math.round(Number(file.Length || 0) / 1024))} KB</span></div><a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.openLink}>Open ↗</a></div>)}</div>}
          <div className={styles.libraryFoot}><span className={styles.liveDot} /> Connected to Kapture Knowledge Repository</div>
        </section>
      </div>
      <footer className={styles.footer}><span>INDOVANCE <b>·</b> KAPTURE KNOWLEDGE MANAGEMENT</span><span>Built for sharing what we know.</span></footer>
    </div>;
  }
}
