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
    destination += `/${this.state.category}`; await this.ensureFolder(destination);
    return destination;
  }
  private async readIndex(): Promise<{ [key: string]: any }> {
    const root = this.props.repositoryFolder.replace(/\/$/, '');
    const url = this.api(`/web/GetFileByServerRelativeUrl('${this.escapeOData(root + '/_KaptureIndex.json')}')/$value`);
    const response = await this.request(url, 'GET');
    if (response.status === 404) return {};
    if (!response.ok) return {};
    try { return JSON.parse(await response.text()); } catch (_) { return {}; }
  }
  private async writeIndex(index: { [key: string]: any }): Promise<void> {
    const root = this.props.repositoryFolder.replace(/\/$/, '');
    const url = this.api(`/web/GetFolderByServerRelativeUrl('${this.escapeOData(root)}')/Files/add(url='_KaptureIndex.json',overwrite=true)`);
    const response = await this.request(url, 'POST', new Blob([JSON.stringify(index)], { type: 'application/json' }));
    if (!response.ok) throw new Error(`The document uploaded, but its searchable description/tag index could not be saved (${response.status}).`);
  }
  private async loadFiles(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const root = this.props.repositoryFolder.replace(/\/$/, '');
      const escapedRoot = this.escapeOData(root);
      const folderCheck = await this.request(this.api(`/web/GetFolderByServerRelativeUrl('${escapedRoot}')?$select=ServerRelativeUrl,Name`), 'GET');
      if (!folderCheck.ok) throw new Error(`The repository folder was not found at '${root}' (${folderCheck.status}). Confirm it exists inside Documents.`);
      let nextUrl: string | undefined = this.api(`/web/lists/getbytitle('Documents')/items?$select=FileLeafRef,FileRef,FileDirRef,File_x0020_Size,Modified,FSObjType&$filter=FSObjType eq 0 and startswith(FileDirRef,'${escapedRoot}')&$orderby=Modified desc&$top=500`);
      const allItems: any[] = [];
      while (nextUrl) {
        const response = await this.request(nextUrl, 'GET');
        if (!response.ok) throw new Error(`Could not read repository files (${response.status}). Check that you have permission to view Documents.`);
        const data: any = await response.json(); allItems.push(...(data.value || []));
        nextUrl = data['@odata.nextLink'] || data['odata.nextLink'] || undefined;
      }
      const index = await this.readIndex();
      const prefix = `${root}/`;
      const files: IKnowledgeFile[] = allItems.filter((item: any) => item.FileRef && item.FileRef.indexOf(prefix) === 0 && item.FileLeafRef !== '_KaptureIndex.json').map((item: any) => {
        const relative = String(item.FileDirRef || '').substring(prefix.length).split('/').filter(Boolean);
        const department = relative[0] || '';
        const subDepartment = department === 'Operations' ? (relative[1] || '') : '';
        const category = department === 'Operations' ? (relative[2] || '') : (relative[1] || '');
        const extra = index[item.FileRef] || {};
        return { Name: item.FileLeafRef, ServerRelativeUrl: item.FileRef, Length: String(item.File_x0020_Size || 0), TimeLastModified: item.Modified, Department: department, SubDepartment: subDepartment, Category: category, Tags: extra.tags || '', Description: extra.description || '' } as IKnowledgeFile;
      }).filter((f: IKnowledgeFile) => DEPARTMENTS.indexOf(f.Department || '') >= 0);
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
      const destination = await this.ensureDestination();
      const safeName = this.escapeOData(file.name);
      const uploadUrl = this.api(`/web/GetFolderByServerRelativeUrl('${this.escapeOData(destination)}')/Files/add(url='${safeName}',overwrite=true)`);
      const uploadResponse = await this.request(uploadUrl, 'POST', file);
      if (!uploadResponse.ok) throw new Error(`Upload failed (${uploadResponse.status}). Check file size, file type and SharePoint edit permissions.`);
      const uploaded: any = await uploadResponse.json();
      const fileUrl = uploaded.ServerRelativeUrl || `${destination}/${file.name}`;
      const index = await this.readIndex();
      index[fileUrl] = { department: this.state.department, subDepartment: this.state.subDepartment, category: this.state.category, description: this.state.description, tags: this.state.tags.join('; '), updated: new Date().toISOString() };
      let indexWarning = '';
      try { await this.writeIndex(index); } catch (e) { indexWarning = ' The file is saved, but its description and tags index could not be updated.'; }
      this.setState({ uploading: false, message: `Successfully uploaded ${file.name} to ${this.state.department}${this.state.subDepartment ? ` / ${this.state.subDepartment}` : ''} / ${this.state.category}.${indexWarning}`, selectedFile: undefined, description: '', tags: [] });
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
          <div className={styles.secureNote}><span>▣</span> Files are stored in SharePoint folders. Access remains governed by your organisation's permissions.</div>
        </section>
        <section className={styles.libraryCard}><div className={styles.libraryHead}><div><span className={styles.kicker}>YOUR CONTENT</span><h3>Repository documents</h3><p>Files stored in the knowledge repository.</p></div><div className={styles.count}>{this.state.files.length}<small>FILES</small></div></div>
          <div className={styles.searchBox}><span>⌕</span><input value={this.state.search} onChange={e => this.setState({ search: e.target.value })} placeholder="Search documents by name…" /></div>
          {this.state.loading ? <div className={styles.empty}>Loading repository…</div> : this.visibleFiles.length === 0 ? <div className={styles.empty}><span>▤</span><strong>{this.state.files.length ? 'No matching documents' : 'Your knowledge library is ready'}</strong><small>{this.state.files.length ? 'Try another search term.' : 'Your uploaded resources will appear here.'}</small></div> : <div className={styles.fileList}>{this.visibleFiles.map(file => <div className={styles.file} key={file.ServerRelativeUrl}><div className={styles.fileIcon}>{file.Name.toLowerCase().match(/\.(mp4|mov|wmv|avi|m4v|webm)$/) ? '▶' : 'DOC'}</div><div className={styles.fileInfo}><strong>{file.Name}</strong><span>Modified {this.formatDate(file.TimeLastModified)} · {Math.max(1, Math.round(Number(file.Length || 0) / 1024))} KB</span></div><a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.openLink}>Open ↗</a></div>)}</div>}
          <div className={styles.libraryFoot}><span className={styles.liveDot} /> Repository contents loaded</div>
        </section>
      </div>
      <section className={styles.analytics}><div className={styles.analyticsHeading}><div><span className={styles.kicker}>REPOSITORY INSIGHTS</span><h2>Knowledge coverage</h2><p>Live counts calculated from each document’s repository folder path.</p></div><div className={styles.totalMetric}><strong>{this.state.files.length}</strong><span>TOTAL DOCUMENTS</span></div></div><div className={styles.analyticsGrid}><div className={styles.analyticsCard}><h3>Documents by department</h3><p>Where organisational knowledge is currently stored</p>{this.departmentStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxDept) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Documents by category</h3><p>Content mix across the repository</p>{this.categoryStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxCat) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Operations capabilities</h3><p>Knowledge coverage across Operations sub-departments</p>{this.operationsStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxOps) * 100}%` }} /></div></div>)}</div></div><div className={styles.analyticsNote}>Department, sub-department and category counts are inferred from folder locations. Descriptions and selected tags are stored in a small application index file in the repository root.</div></section>
      <footer className={styles.footer}><span>INDOVANCE <b>·</b> KNOWLEDGE OPERATIONS</span><span>Built for sharing what we know.</span></footer>
    </div>;
  }
}
