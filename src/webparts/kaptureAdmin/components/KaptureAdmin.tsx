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
  selectedFile?: File; selectedFiles: File[]; failedFiles: File[]; uploadProgress: string; page: number; pageSize: number; previewFile?: IKnowledgeFile; department: string; subDepartment: string; category: string; description: string; tags: string[]; search: string; filterDepartment: string; filterCategory: string; manifest: { [name: string]: { department: string; subDepartment: string; category: string; description: string; tags: string } };
}

export default class KaptureAdmin extends React.Component<IKaptureAdminProps, IState> {
  constructor(props: IKaptureAdminProps) {
    super(props);
    this.state = { files: [], loading: true, uploading: false, message: '', error: '', selectedFiles: [], failedFiles: [], uploadProgress: '', page: 1, pageSize: 10, department: '', subDepartment: '', category: '', description: '', tags: [], search: '', filterDepartment: '', filterCategory: '', manifest: {} };
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
  private async readPagedCollection(url: string): Promise<any[]> {
    const results: any[] = [];
    let nextUrl: string | undefined = url;
    while (nextUrl) {
      const response = await this.request(nextUrl, 'GET');
      if (!response.ok) throw new Error(`SharePoint returned ${response.status} while reading a repository folder. Confirm you have at least read access to the Kapture Knowledge Repository.`);
      const data: any = await response.json();
      results.push(...(data.value || []));
      nextUrl = data['@odata.nextLink'] || data['odata.nextLink'] || undefined;
    }
    return results;
  }
  private async collectFolderFiles(folderPath: string, output: any[]): Promise<void> {
    const escaped = this.escapeOData(folderPath);
    const base = `/web/GetFolderByServerRelativeUrl('${escaped}')`;
    const files = await this.readPagedCollection(this.api(`${base}/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified`));
    files.forEach((file: any) => { if (file.Name !== '_KaptureIndex.json') output.push(file); });
    const folders = await this.readPagedCollection(this.api(`${base}/Folders?$select=Name,ServerRelativeUrl`));
    for (const folder of folders) {
      if (folder.Name && folder.Name !== 'Forms' && folder.ServerRelativeUrl) {
        await this.collectFolderFiles(folder.ServerRelativeUrl, output);
      }
    }
  }
  private async loadFiles(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const root = this.props.repositoryFolder.replace(/\/$/, '');
      const escapedRoot = this.escapeOData(root);
      const folderCheck = await this.request(this.api(`/web/GetFolderByServerRelativeUrl('${escapedRoot}')?$select=ServerRelativeUrl,Name`), 'GET');
      if (!folderCheck.ok) throw new Error(`The repository folder was not found at '${root}' (${folderCheck.status}). Confirm it exists inside the Documents library and that you can open it.`);
      // Enumerate the repository through the folder REST API. The previous
      // list-items query used startswith(FileDirRef, ...) and caused HTTP 400
      // in this tenant. Folder enumeration avoids that unsupported filter.
      const allItems: any[] = [];
      await this.collectFolderFiles(root, allItems);
      const index = await this.readIndex();
      const prefix = `${root}/`;
      const files: IKnowledgeFile[] = allItems.filter((item: any) => item.ServerRelativeUrl && item.ServerRelativeUrl.indexOf(prefix) === 0).map((item: any) => {
        const parent = String(item.ServerRelativeUrl).substring(0, String(item.ServerRelativeUrl).lastIndexOf('/'));
        const relative = parent.substring(prefix.length).split('/').filter(Boolean);
        const department = relative[0] || '';
        const subDepartment = department === 'Operations' ? (relative[1] || '') : '';
        const category = department === 'Operations' ? (relative[2] || '') : (relative[1] || '');
        const extra = index[item.ServerRelativeUrl] || {};
        return { Name: item.Name, ServerRelativeUrl: item.ServerRelativeUrl, Length: String(item.Length || 0), TimeLastModified: item.TimeLastModified || '', Department: department, SubDepartment: subDepartment, Category: category, Tags: extra.tags || '', Description: extra.description || '' } as IKnowledgeFile;
      }).filter((f: IKnowledgeFile) => DEPARTMENTS.indexOf(f.Department || '') >= 0);
      this.setState({ files, loading: false });
    } catch (e) { this.setState({ loading: false, error: e instanceof Error ? e.message : 'Unable to load repository documents.' }); }
  }
  private onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => { const files = event.target.files ? Array.prototype.slice.call(event.target.files) as File[] : []; this.setState({ selectedFiles: files, selectedFile: files[0], failedFiles: [], message: '', error: '' }); };
  private parseCsv(text: string): string[][] {
    const rows: string[][] = []; let row: string[] = []; let cell = ''; let quoted = false;
    for (let i = 0; i < text.length; i++) { const ch = text[i]; if (ch === '"' && quoted && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') quoted = !quoted; else if (ch === ',' && !quoted) { row.push(cell); cell = ''; } else if ((ch === '\n' || ch === '\r') && !quoted) { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); if (row.some(v => v.trim())) rows.push(row); row = []; cell = ''; } else cell += ch; }
    row.push(cell); if (row.some(v => v.trim())) rows.push(row); return rows;
  }
  private onManifestChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files && event.target.files[0]; if (!file) return;
    try { const rows = this.parseCsv(await file.text()); if (rows.length < 2) throw new Error('The manifest has no document rows.'); const headers = rows[0].map(h => h.trim().toLowerCase()); const col = (name: string) => headers.indexOf(name.toLowerCase());
      if (col('filename') < 0 || col('department') < 0 || col('category') < 0) throw new Error('Manifest must include FileName, Department and Category columns.');
      const manifest: any = {}; rows.slice(1).forEach(r => { const name = (r[col('filename')] || '').trim(); if (name) manifest[name] = { department: (r[col('department')] || '').trim(), subDepartment: col('subdepartment') >= 0 ? (r[col('subdepartment')] || '').trim() : '', category: (r[col('category')] || '').trim(), description: col('description') >= 0 ? (r[col('description')] || '').trim() : '', tags: col('tags') >= 0 ? (r[col('tags')] || '').trim() : '' }; });
      this.setState({ manifest, message: `Loaded metadata for ${Object.keys(manifest).length} files from the manifest.`, error: '' });
    } catch (e) { this.setState({ error: e instanceof Error ? e.message : 'Could not read manifest.' }); }
    event.target.value = '';
  };
  private downloadManifestTemplate = (): void => {
    const csv = 'FileName,Department,SubDepartment,Category,Description,Tags\r\n"Example-SOP.pdf",Mechanical,,SOP,"Purpose and scope","Process; Quality"\r\n"Example-Training.pdf",Operations,"Building Services","Training Material","Training reference","Training; Reference"';
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'Kapture-Bulk-Upload-Manifest.csv'; a.click(); URL.revokeObjectURL(url);
  };
  private async uploadBatch(): Promise<void> {
    const files = this.state.failedFiles.length ? this.state.failedFiles : this.state.selectedFiles;
    if (!files.length) { this.setState({ error: 'Choose one or more files before uploading.' }); return; }
    if (!Object.keys(this.state.manifest).length && (!this.state.department || !this.state.category || (this.state.department === 'Operations' && !this.state.subDepartment))) { this.setState({ error: 'Choose department, Operations sub-department (if applicable) and category, or upload a completed CSV manifest.' }); return; }
    const failures: File[] = []; let metadataWarnings = 0; let completed = 0; this.setState({ uploading: true, error: '', message: '', uploadProgress: `Preparing ${files.length} files…` });
    for (const file of files) {
      const meta = this.state.manifest[file.name] || { department: this.state.department, subDepartment: this.state.subDepartment, category: this.state.category, description: this.state.description, tags: this.state.tags.join('; ') };
      if (!meta.department || DEPARTMENTS.indexOf(meta.department) < 0 || !meta.category || CATEGORIES.indexOf(meta.category) < 0 || (meta.department === 'Operations' && OPERATIONS.indexOf(meta.subDepartment) < 0)) { failures.push(file); completed++; this.setState({ uploadProgress: `${completed} of ${files.length} checked · ${file.name} needs valid manifest metadata` }); continue; }
      try {
        const root = this.props.repositoryFolder.replace(/\/$/, ''); let destination = root;
        await this.ensureFolder(destination); if (meta.department === 'Operations') { destination += '/Operations'; await this.ensureFolder(destination); destination += `/${meta.subDepartment}`; await this.ensureFolder(destination); } else { destination += `/${meta.department}`; await this.ensureFolder(destination); }
        destination += `/${meta.category}`; await this.ensureFolder(destination);
        const uploadUrl = this.api(`/web/GetFolderByServerRelativeUrl('${this.escapeOData(destination)}')/Files/add(url='${this.escapeOData(file.name)}',overwrite=false)`);
        const response = await this.request(uploadUrl, 'POST', file); if (!response.ok) throw new Error(`SharePoint upload returned ${response.status}`);
        const uploaded: any = await response.json(); const fileUrl = uploaded.ServerRelativeUrl || `${destination}/${file.name}`;
        // Read-modify-write for each file, sequentially, to avoid overwriting another file's metadata within this batch.
        const index = await this.readIndex(); index[fileUrl] = { department: meta.department, subDepartment: meta.subDepartment, category: meta.category, description: meta.description || '', tags: meta.tags || '', updated: new Date().toISOString() }; try { await this.writeIndex(index); } catch (_) { metadataWarnings++; }
      } catch (_) { failures.push(file); }
      completed++; this.setState({ uploadProgress: `${completed} of ${files.length} processed · ${failures.length} failed` });
    }
    this.setState({ uploading: false, failedFiles: failures, selectedFiles: failures, selectedFile: failures[0], message: `${files.length - failures.length} of ${files.length} files uploaded successfully.${failures.length ? ` ${failures.length} failed; use Retry failed files after checking their metadata or permissions.` : ''}${metadataWarnings ? ` Metadata index could not be updated for ${metadataWarnings} uploaded file(s); the files themselves are in SharePoint.` : ''}`, error: failures.length ? 'Some files could not be uploaded. Successful files have been retained and will not be retried.' : '', uploadProgress: '' });
    await this.loadFiles();
  }
  private async upload(): Promise<void> { return this.uploadBatch(); }
  private openPreview = (file: IKnowledgeFile): void => { this.setState({ previewFile: file }); };
  private formatDate(value: string): string { return value ? new Date(value).toLocaleDateString() : '—'; }
  private toggleTag = (tag: string): void => { this.setState(state => ({ tags: state.tags.indexOf(tag) >= 0 ? state.tags.filter(t => t !== tag) : [...state.tags, tag] })); };
  private get departmentStats(): { name: string; count: number }[] { return DEPARTMENTS.map(name => ({ name, count: this.state.files.filter(f => f.Department === name).length })).sort((a,b) => b.count-a.count); }
  private get categoryStats(): { name: string; count: number }[] { return CATEGORIES.map(name => ({ name, count: this.state.files.filter(f => f.Category === name).length })).sort((a,b) => b.count-a.count); }
  private get operationsStats(): { name: string; count: number }[] { return OPERATIONS.map(name => ({ name, count: this.state.files.filter(f => f.Department === 'Operations' && f.SubDepartment === name).length })).sort((a,b) => b.count-a.count); }
  private get visibleFiles(): IKnowledgeFile[] { const term = this.state.search.toLowerCase(); return this.state.files.filter(f => f.Name.toLowerCase().indexOf(term) >= 0 && (!this.state.filterDepartment || f.Department === this.state.filterDepartment) && (!this.state.filterCategory || f.Category === this.state.filterCategory)); }
  public render(): React.ReactElement<IKaptureAdminProps> {
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
          <label className={styles.dropzone}><input id="kapture-admin-file" type="file" multiple accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.mp4,.mov,.wmv,.avi,.m4v,.webm" onChange={this.onFileChange} /><span className={styles.uploadGlyph}>⇧</span><strong>{this.state.selectedFiles.length ? `${this.state.selectedFiles.length} files selected` : 'Choose files to upload'}</strong><small>{this.state.selectedFiles.length ? `${this.state.selectedFiles.slice(0, 3).map(f => f.name).join(', ')}${this.state.selectedFiles.length > 3 ? ` and ${this.state.selectedFiles.length - 3} more` : ''}` : 'Select multiple documents or videos in one batch'}</small><em>Browse files</em></label>
          <div className={styles.formGrid}><label className={styles.field}>Department<select value={this.state.department} onChange={e => this.setState({ department: e.target.value, subDepartment: '' })}><option value="">Select department</option>{DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>
            {this.state.department === 'Operations' && <label className={styles.field}>Operations sub-department<select value={this.state.subDepartment} onChange={e => this.setState({ subDepartment: e.target.value })}><option value="">Select sub-department</option>{OPERATIONS.map(d => <option key={d} value={d}>{d}</option>)}</select></label>}
            <label className={styles.field}>Document category<select value={this.state.category} onChange={e => this.setState({ category: e.target.value })}><option value="">Select category</option>{CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}</select></label>
            <label className={`${styles.field} ${styles.full}`}>Description <textarea rows={2} value={this.state.description} onChange={e => this.setState({ description: e.target.value })} placeholder="What is this resource about? (optional)" /></label>
            <div className={`${styles.field} ${styles.full}`}><span className={styles.fieldLabel}>Suggested tags <small>Select all that apply</small></span><div className={styles.tagPicker}>{TAG_OPTIONS.map(tag => <button type="button" key={tag} className={`${styles.tagOption} ${this.state.tags.indexOf(tag) >= 0 ? styles.tagSelected : ''}`} onClick={() => this.toggleTag(tag)}>{this.state.tags.indexOf(tag) >= 0 ? '✓ ' : '+ '}{tag}</button>)}</div></div></div>
          <div style={{display:'flex',gap:10,flexWrap:'wrap',margin:'12px 0'}}><button type="button" className={styles.refresh} onClick={this.downloadManifestTemplate}>Download Excel-compatible CSV template</button><label className={styles.refresh} style={{cursor:'pointer'}}>Import completed manifest<input type="file" accept=".csv,text/csv" onChange={this.onManifestChange} style={{display:'none'}} /></label></div><small style={{display:'block',color:'#64748b',marginBottom:10}}>{Object.keys(this.state.manifest).length ? `Manifest loaded: ${Object.keys(this.state.manifest).length} filename mappings. Filenames must match exactly.` : 'Use common metadata for the whole batch, or import a CSV manifest (opens in Excel) for different metadata per file.'}</small><button className={styles.primary} disabled={this.state.uploading} onClick={() => this.upload()}>{this.state.uploading ? <><span className={styles.spinner} /> Uploading batch…</> : this.state.failedFiles.length ? `Retry ${this.state.failedFiles.length} failed files →` : `Upload ${this.state.selectedFiles.length || ''} files to repository →`}</button>
          {this.state.uploadProgress && <div style={{marginTop:12,padding:12,background:'#f1f5f9',borderRadius:8,color:'#334155'}}>{this.state.uploadProgress}</div>}<div className={styles.secureNote}><span>▣</span> Files are stored in SharePoint folders. Access remains governed by your organisation's permissions.</div>
        </section>
        <section className={styles.libraryCard}><div className={styles.libraryHead}><div><span className={styles.kicker}>YOUR CONTENT</span><h3>Repository documents</h3><p>Files stored in the knowledge repository.</p></div><div className={styles.count}>{this.state.files.length}<small>FILES</small></div></div>
          <div className={styles.searchBox}><span>⌕</span><input value={this.state.search} onChange={e => this.setState({ search: e.target.value, page: 1 })} placeholder="Search documents by name…" /></div><div style={{display:'flex',gap:8,margin:'10px 0',flexWrap:'wrap'}}><select aria-label="Filter department" value={this.state.filterDepartment} onChange={e=>this.setState({filterDepartment:e.target.value,page:1})}><option value="">All departments</option>{DEPARTMENTS.map(d=><option key={d} value={d}>{d}</option>)}</select><select aria-label="Filter category" value={this.state.filterCategory} onChange={e=>this.setState({filterCategory:e.target.value,page:1})}><option value="">All categories</option>{CATEGORIES.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
          {this.state.loading ? <div className={styles.empty}>Loading repository…</div> : this.visibleFiles.length === 0 ? <div className={styles.empty}><span>▤</span><strong>{this.state.files.length ? 'No matching documents' : 'Your knowledge library is ready'}</strong><small>{this.state.files.length ? 'Try another search term.' : 'Your uploaded resources will appear here.'}</small></div> : <div className={styles.fileList}>{this.visibleFiles.slice((this.state.page-1)*this.state.pageSize,this.state.page*this.state.pageSize).map(file => <div className={styles.file} key={file.ServerRelativeUrl}><div className={styles.fileIcon}>{file.Name.toLowerCase().match(/\.(mp4|mov|wmv|avi|m4v|webm)$/) ? '▶' : 'DOC'}</div><div className={styles.fileInfo}><strong>{file.Name}</strong><span>Modified {this.formatDate(file.TimeLastModified)} · {Math.max(1, Math.round(Number(file.Length || 0) / 1024))} KB</span></div><button type="button" onClick={()=>this.openPreview(file)} className={styles.openLink}>Preview ↗</button></div>)}</div>}
          {this.visibleFiles.length > this.state.pageSize && <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'12px 0'}}><span style={{fontSize:12,color:'#64748b'}}>Showing {(this.state.page-1)*this.state.pageSize+1}–{Math.min(this.state.page*this.state.pageSize,this.visibleFiles.length)} of {this.visibleFiles.length}</span><div style={{display:'flex',gap:8}}><button className={styles.refresh} disabled={this.state.page===1} onClick={()=>this.setState({page:this.state.page-1})}>Previous</button><button className={styles.refresh} disabled={this.state.page*this.state.pageSize>=this.visibleFiles.length} onClick={()=>this.setState({page:this.state.page+1})}>Next</button></div></div>}<div className={styles.libraryFoot}><span className={styles.liveDot} /> Repository contents loaded</div>
        </section>
      </div>
      <section className={styles.analytics}><div className={styles.analyticsHeading}><div><span className={styles.kicker}>REPOSITORY INSIGHTS</span><h2>Knowledge coverage</h2><p>Live counts calculated from each document’s repository folder path.</p></div><div className={styles.totalMetric}><strong>{this.state.files.length}</strong><span>TOTAL DOCUMENTS</span></div></div><div className={styles.analyticsGrid}><div className={styles.analyticsCard}><h3>Documents by department</h3><p>Where organisational knowledge is currently stored</p>{this.departmentStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxDept) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Documents by category</h3><p>Content mix across the repository</p>{this.categoryStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxCat) * 100}%` }} /></div></div>)}</div><div className={styles.analyticsCard}><h3>Operations capabilities</h3><p>Knowledge coverage across Operations sub-departments</p>{this.operationsStats.map(item => <div className={styles.statRow} key={item.name}><div><span>{item.name}</span><b>{item.count}</b></div><div className={styles.barTrack}><i style={{ width: `${(item.count / maxOps) * 100}%` }} /></div></div>)}</div></div><div className={styles.analyticsNote}>Department, sub-department and category counts are inferred from folder locations. Descriptions and selected tags are stored in a small application index file in the repository root.</div></section>
      {this.state.previewFile && <div role="dialog" aria-modal="true" aria-label="Document preview" style={{position:'fixed',inset:0,zIndex:10000,background:'rgba(15,23,42,.72)',display:'flex',alignItems:'center',justifyContent:'center',padding:20}}><div style={{background:'#fff',width:'min(1200px,96vw)',height:'min(850px,92vh)',borderRadius:14,overflow:'hidden',display:'flex',flexDirection:'column'}}><div style={{padding:'12px 18px',display:'flex',justifyContent:'space-between',alignItems:'center',borderBottom:'1px solid #e2e8f0'}}><strong>{this.state.previewFile.Name}</strong><button className={styles.refresh} onClick={()=>this.setState({previewFile:undefined})}>Close preview ✕</button></div><iframe title={this.state.previewFile.Name} src={`${this.state.previewFile.ServerRelativeUrl}?web=1`} style={{border:0,width:'100%',height:'100%'}} /></div></div>}<footer className={styles.footer}><span>INDOVANCE <b>·</b> KNOWLEDGE OPERATIONS</span><span>Built for sharing what we know.</span></footer>
    </div>;
  }
}
