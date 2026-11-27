import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IKaptureAdminProps, IKnowledgeFile } from './IKaptureAdminProps';
import styles from './KaptureAdmin.module.scss';

const DEPARTMENTS: string[] = ['HR', 'Finance', 'Marketing', 'Sales', 'Training and Development', 'IT', 'BES', 'AI', 'Technology', 'Operations'];
const OPERATIONS: string[] = ['Land Survey', 'Land Development', 'Structural', 'Visual Engineering', 'Mechanical', 'Building Services', 'Construction Engineering', 'Water and Drainage', 'Energy and Utilities', 'Heavy Civil and Infra'];
const CATEGORIES: string[] = ['SOP', 'Policy', 'Standards', 'Training Material', 'Video & Recorded Sessions', 'Work Instructions', 'Guidelines', 'Templates & Forms', 'Project Reference', 'Technical Documentation', 'Other'];
const TAG_OPTIONS: string[] = ['Process', 'Project', 'Client', 'Internal', 'Reference', 'Checklist', 'Guide', 'Template', 'CAD', 'BIM', 'Revit', 'AutoCAD', 'Structural', 'Mechanical', 'Quality', 'Delivery', 'Onboarding', 'Training', 'Best Practice', 'Lessons Learned', 'Troubleshooting', 'Safety', 'Compliance', 'Design', 'Engineering'];

interface IState {
  files: IKnowledgeFile[]; loading: boolean; uploading: boolean; message: string; error: string;
  selectedFile?: File; department: string; subDepartment: string; category: string; description: string; tags: string[]; search: string;
}

export default class KaptureAdmin extends React.Component<IKaptureAdminProps, IState> {
  constructor(props: IKaptureAdminProps) {
    super(props);
    this.state = { files: [], loading: true, uploading: false, message: '', error: '', department: '', subDepartment: '', category: '', description: '', tags: [], search: '' };
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
    // Metadata columns belong to the parent Documents library, not to the repository folder.
    // Query the Fields collection with an OData filter rather than the single-field
    // getbyinternalnameortitle route, which can return HTTP 400 in some SharePoint tenants.
    const fields = [
      { name: 'KaptureDepartment', title: 'Kapture Department', type: 'Text' },
      { name: 'KaptureSubDepartment', title: 'Kapture Sub-department', type: 'Text' },
      { name: 'KaptureCategory', title: 'Kapture Category', type: 'Text' },
      { name: 'KaptureDescription', title: 'Kapture Description', type: 'Note' },
      { name: 'KaptureTags', title: 'Kapture Tags', type: 'Note' }
    ];
    const listApi = this.api("/web/lists/getbytitle('Documents')");
    const listCheck = await this.request(listApi + '?$select=Id,Title', 'GET');
    if (!listCheck.ok) {
      const detail = await listCheck.text().catch(() => '');
      throw new Error(`SharePoint could not open the 'Documents' library (${listCheck.status}). ${detail.substring(0, 300)}`);
    }
    for (const field of fields) {
      const filter = encodeURIComponent(`InternalName eq '${field.name}'`);
      const fieldUrl = `${listApi}/fields?$select=Id,InternalName,Title&$filter=${filter}`;
      const check = await this.request(fieldUrl, 'GET');
      if (!check.ok) {
        const detail = await check.text().catch(() => '');
        throw new Error(`Could not inspect the '${field.title}' column (${check.status}). SharePoint response: ${detail.substring(0, 350)}`);
      }
      const fieldData: any = await check.json();
      if ((fieldData.value || []).some((item: any) => item.InternalName === field.name)) continue;

      const xml = `<Field Type="${field.type}" DisplayName="${field.title}" Name="${field.name}" StaticName="${field.name}" Group="Kapture Knowledge" />`;
      const add = await this.request(`${listApi}/fields/addFieldAsXml`, 'POST', JSON.stringify({
        parameters: {
          __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
          SchemaXml: xml,
          Options: 0
        }
      }));
      if (!add.ok) {
        // Recheck through the same Fields collection in case another request created it.
        const confirm = await this.request(fieldUrl, 'GET');
        if (confirm.ok) {
          const confirmData: any = await confirm.json();
          if ((confirmData.value || []).some((item: any) => item.InternalName === field.name)) continue;
        }
        const detail = await add.text().catch(() => '');
        if (add.status === 403) {
          throw new Error(`The '${field.title}' column is missing and your account cannot create library columns (403). Ask a SharePoint administrator to create the five Kapture metadata columns once.`);
        }
        throw new Error(`SharePoint could not create the '${field.title}' column (${add.status}). SharePoint response: ${detail.substring(0, 350)}`);
      }
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
      const root = this.props.repositoryFolder.replace(/\/$/, '');
      const escapedRoot = this.escapeOData(root);
      await this.ensureMetadataFields();
      const folderCheck = await this.request(this.api(`/web/GetFolderByServerRelativeUrl('${escapedRoot}')?$select=ServerRelativeUrl,Name`), 'GET');
      if (!folderCheck.ok) {
        throw new Error(`The repository folder was not found at '${root}' (${folderCheck.status}). Confirm the folder exists inside the Documents library.`);
      }
      let nextUrl: string | undefined = this.api(`/web/lists/getbytitle('Documents')/items?$select=File/Name,File/ServerRelativeUrl,File/Length,Modified,KaptureDepartment,KaptureSubDepartment,KaptureCategory,KaptureTags&$expand=File&$filter=FSObjType eq 0 and startswith(FileDirRef,'${escapedRoot}')&$orderby=Modified desc&$top=500`);
      const allItems: any[] = [];
      while (nextUrl) {
        const response: SPHttpClientResponse = await this.request(nextUrl, 'GET');
        if (!response.ok) throw new Error(`Could not read files from the Kapture Knowledge Repository (${response.status}). Check the Documents library permissions and metadata columns.`);
        const data: any = await response.json();
        allItems.push(...(data.value || []));
        nextUrl = data['@odata.nextLink'] || data['odata.nextLink'] || undefined;
      }
      const rootPrefix = `${root}/`;
      const files: IKnowledgeFile[] = allItems.map((item: any) => ({ Name: item.File.Name, ServerRelativeUrl: item.File.ServerRelativeUrl, Length: String(item.File.Length || 0), TimeLastModified: item.Modified, Department: item.KaptureDepartment || '', SubDepartment: item.KaptureSubDepartment || '', Category: item.KaptureCategory || '', Tags: item.KaptureTags || '' })).filter((file: IKnowledgeFile) => file.ServerRelativeUrl && file.ServerRelativeUrl.indexOf(rootPrefix) === 0);
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
        { FieldName: 'KaptureTags', FieldValue: this.state.tags.join('; ') }
      ];
      const metadataResponse = await this.request(this.api(`/web/GetFileByServerRelativeUrl('${this.escapeOData(fileUrl)}')/ListItemAllFields/ValidateUpdateListItem`), 'POST', JSON.stringify({ formValues: metadata, bNewDocumentUpdate: false }));
      if (!metadataResponse.ok) throw new Error('The file was uploaded, but its metadata could not be saved. Please ask the SharePoint administrator to check the Kapture metadata columns.');
      this.setState({ uploading: false, message: `Successfully uploaded ${file.name} to ${this.state.department}${this.state.subDepartment ? ` / ${this.state.subDepartment}` : ''}.`, selectedFile: undefined, description: '', tags: [] });
      const input = document.getElementById('kapture-admin-file') as HTMLInputElement; if (input) input.value = '';
      await this.loadFiles();
    } catch (e) { this.setState({ uploading: false, error: e instanceof Error ? e.message : 'Upload failed.' }); }
  }
  private formatDate(value: string): string { return value ? new Date(value).toLocaleDateString() : '—'; }
  private toggleTag = (tag: string): void => { this.setState(state => ({ tags: state.tags.indexOf(tag) >= 0 ? state.tags.filter(t => t !== tag) : [...state.tags, tag] })); };
  private get departmentStats(): { name: string; count: number }[] { return DEPARTMENTS.map(name => ({ name, count: this.state.files.filter(f => f.Department === name).length })).sort((a,b) => b.count-a.count); }
  private get categoryStats(): { name: string; count: number }[] { return CATEGORIES.map(name => ({ name, count: this.state.files.filter(f => f.Category === name).length })).sort((a,b) => b.count-a.count); }
  private get operationsStats(): { name: string; count: number }[] { return OPERATIONS.map(name => ({ name, count: this.state.files.filter(f => f.Department === 'Operations' && f.SubDepartment === name).length })).sort((a,b) => b.count-a.count); }
  private get visibleFiles(): IKnowledgeFile[] { const term = this.state.search.toLowerCase(); return this.state.files.filter(f => f.Name.toLowerCase().indexOf(term) >= 0); }
  public render(): React.ReactElement<IKaptureAdminProps> {
    const isVideo = !!this.state.selectedFile && this.state.selectedFile.type.indexOf('video/') === 0;
    const maxDept = Math.max(1, ...this.departmentStats.map(x => x.count));
    const maxCat = Math.max(1, ...this.categoryStats.map(x => x.count));
    const maxOps = Math.max(1, ...this.operationsStats.map(x => x.count));
    return <div className={styles.root}>
      <div className={styles.topbar}><div className={styles.brandText}><strong className={styles.indovanceWordmark}>INDOVANCE</strong><span>KNOWLEDGE OPERATIONS</span></div><div className={styles.topSpacer} /><div className={styles.profile}><span className={styles.avatar}>{(this.props.displayName || 'A').charAt(0).toUpperCase()}</span><span><small>ADMINISTRATOR</small><strong>{this.props.displayName}</strong></span></div></div>
      <section className={styles.hero}><div className={styles.heroCopy}><div className={styles.eyebrow}>INDOVANCE · KNOWLEDGE OPERATIONS</div><h1>Indovance Knowledge Hub</h1><p>One place to capture, organise and share the knowledge that powers Indovance.</p></div><div className={styles.heroArt}><div className={styles.heroStamp}><strong>INDOVANCE</strong><span>KNOWLEDGE HUB</span><i /></div></div></section>
      <div className={styles.sectionTitle}><div><span className={styles.kicker}>ADMIN WORKSPACE</span><h2>Knowledge operations</h2><p>Upload resources, monitor coverage and manage the knowledge repository.</p></div><button className={styles.refresh} onClick={() => this.loadFiles()}>↻ Refresh repository</button></div>
      {this.state.message && <div className={styles.success}><b>✓</b>{this.state.message}</div>}{this.state.error && <div className={styles.error}><b>!</b>{this.state.error}</div>}
      <div className={styles.layout}>
        <section className={styles.uploadCard}><div className={styles.cardHead}><div className={styles.iconBadge}><span>↑</span></div><div><h3>Add to the knowledge base</h3><p>Share a useful resource with your colleagues.</p></div></div>
          <label className={styles.dropzone}><input id="kapture-admin-file" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.mp4,.mov,.wmv,.avi,.m4v,.webm" onChange={this.onFileChange} /><span className={styles.uploadGlyph}>⇧</span><strong>{this.state.selectedFile ? this.state.selectedFile.name : 'Choose a file to upload'}</strong><small>{this.state.selectedFile ? `${(this.state.selectedFile.size / (1024 * 1024)).toFixed(2)} MB · ${isVideo ? 'Video' : 'Document'}` : 'Drag-and-drop is coming soon · Browse your device'}</small><em>Browse files</em></label>
          <div className={styles.formGrid}><label className={styles.field}>Department<select value={this.state.department} onChange={e => this.setState({ department: e.target.value, subDepartment: '' })}><option value="">Select department</option>{DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>
            {this.state.department === 'Operations' && <label className={styles.field}>Operations sub-department<select value={this.state.subDepartment} onChange={e => this.setState({ subDepartment: e.target.value })}><option value="">Select sub-department</option>{OPERATIONS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>}
            <label className={styles.field}>Document category<select value={this.state.category} onChange={e => this.setState({ category: e.target.value })}><option value="">Select category</option>{CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}</select></label>
            <label className={`${styles.field} ${styles.full}`}>Description <textarea rows={2} value={this.state.description} onChange={e => this.setState({ description: e.target.value })} placeholder="What is this resource about? (optional)" /></label>
            <div className={`${styles.field} ${styles.full}`}><span className={styles.fieldLabel}>Suggested tags <small>Select all that apply</small></span><div className={styles.tagPicker}>{TAG_OPTIONS.map(tag => <button type="button" key={tag} className={`${styles.tagOption} ${this.state.tags.indexOf(tag) >= 0 ? styles.tagSelected : ''}`} onClick={() => this.toggleTag(tag)}>{this.state.tags.indexOf(tag) >= 0 ? '✓ ' : '+ '}{tag}</button>)}</div></div></div>
          <button className={styles.primary} disabled={this.state.uploading} onClick={() => this.upload()}>{this.state.uploading ? <><span className={styles.spinner} /> Uploading resource…</> : 'Upload to knowledge repository  →'}</button>
          <div className={styles.secureNote}><span>▣</span> Files are stored in SharePoint. Access remains governed by your organisation's permissions.</div>
        </section>
        <section className={styles.libraryCard}><div className={styles.libraryHead}><div><span className={styles.kicker}>YOUR CONTENT</span><h3>Repository documents</h3><p>Files stored in the knowledge repository.</p></div><div className={styles.count}>{this.state.files.length}<small>FILES</small></div></div>
          <div className={styles.searchBox}><span>⌕</span><input value={this.state.search} onChange={e => this.setState({ search: e.target.value })} placeholder="Search documents by name…" /></div>
          {this.state.loading ? <div className={styles.empty}>Loading repository…</div> : this.visibleFiles.length === 0 ? <div className={styles.empty}><span>▤</span><strong>{this.state.files.length ? 'No matching documents' : 'Your knowledge library is ready'}</strong><small>{this.state.files.length ? 'Try another search term.' : 'Your uploaded resources will appear here.'}</small></div> : <div className={styles.fileList}>{this.visibleFiles.map(file => <div className={styles.file} key={file.ServerRelativeUrl}><div className={styles.fileIcon}>{file.Name.toLowerCase().match(/\.(mp4|mov|wmv|avi|m4v|webm)$/) ? '▶' : 'DOC'}</div><div className={styles.fileInfo}><strong>{file.Name}</strong><span>Modified {this.formatDate(file.TimeLastModified)} · {Math.max(1, Math.round(Number(file.Length || 0) / 1024))} KB</span></div><a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.openLink}>Open ↗</a></div>)}</div>}
          <div className={styles.libraryFoot}><span className={styles.liveDot} /> Repository contents loaded</div>
        </section>
      </div>
      <section className={styles.analytics}><div className={styles.analyticsHeading}><div><span className={styles.kicker}>REPOSITORY INSIGHTS</span><h2>Knowledge coverage</h2><p>Live counts based on documents and their saved SharePoint metadata.</p></div><div className={styles.totalMetric}><strong>{this.state.files.length}</strong><span>TOTAL DOCUMENTS</span></div></div><div className={styles.analyticsGrid}><div className={styles.analyticsCard}><h3>Documents by department</h3><p>Where organisational knowledge is currently stored</p>{this.departmentStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxDept) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Documents by category</h3><p>Content mix across the repository</p>{this.categoryStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxCat) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Operations capabilities</h3><p>Knowledge coverage across Operations sub-departments</p>{this.operationsStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxOps) * 100}%` }} /></div></div>)}</div></div><div className={styles.analyticsNote}>Counts are calculated from all repository items returned through SharePoint paging. Documents without department or category metadata are included in the total but not assigned to a breakdown.</div></section>
      <footer className={styles.footer}><span>INDOVANCE <b>·</b> KNOWLEDGE OPERATIONS</span><span>Built for sharing what we know.</span></footer>
    </div>;
  }
}
