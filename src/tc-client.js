// tc-client.js — minimal Teamcenter 2412 SOA (JSON/REST) + FMS client.
//
// Talks to the Teamcenter web tier through the add-in's own server, which proxies
//   /tc/*   -> Teamcenter web tier  (JsonRestServices)
//   /fms/*  -> FSC (File Server Cache)
// so the browser stays same-origin (no CORS) and the TC session cookie (JSESSIONID)
// is handled by the browser automatically.
//
// Works in the browser and in Node 18+ (tests inject a fetch with a cookie jar).

const NULL_TAG = 'AAAAAAAAAAAAAA';

export class TcError extends Error {
  constructor(message, details) {
    super(message);
    this.name = 'TcError';
    this.details = details;
  }
}

/** Teamcenter serialises maps as [[keys],[values]]; accept that or a plain object. */
export function mapEntries(m) {
  if (!m) return [];
  if (Array.isArray(m) && m.length === 2 && Array.isArray(m[0]) && Array.isArray(m[1])) {
    return m[0].map((k, i) => [k, m[1][i]]);
  }
  if (typeof m === 'object') return Object.entries(m);
  return [];
}

export const ref = (o) => (o ? { uid: o.uid, type: o.type || o.className || 'unknownType' } : { uid: NULL_TAG, type: 'unknownType' });

export class TcClient {
  /**
   * @param {object} opts
   * @param {string} [opts.baseUrl]  origin of the add-in server ('' = same origin)
   * @param {string} [opts.tcPath]   path of the TC web tier on that server (default '/tc')
   * @param {string} [opts.fmsPath]  path of the FSC proxy on that server (default '/fms')
   * @param {Function} [opts.fetch]  fetch implementation (tests)
   */
  constructor(opts = {}) {
    this.baseUrl = opts.baseUrl || '';
    this.tcPath = opts.tcPath || '/tc';
    this.fmsPath = opts.fmsPath || '/fms';
    this._fetch = opts.fetch || ((...a) => fetch(...a));
    this.objects = new Map(); // uid -> model object (client-side cache of last props seen)
    this.session = null;      // { user, group, role }
  }

  // ---------------------------------------------------------------- core call
  async call(service, operation, body = {}, policy = {}) {
    const url = `${this.baseUrl}${this.tcPath}/JsonRestServices/${service}/${operation}`;
    const envelope = {
      header: {
        state: { clientVersion: '10000.1.2', stateless: true, unloadObjects: true, enableServerStateHeaders: true, formatProperties: true, clientID: 'EPCM-OfficeAddin' },
        policy,
      },
      body,
    };
    const res = await this._fetch(url, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(envelope),
    });
    const text = await res.text();
    if (!res.ok) throw new TcError(`${service}/${operation}: HTTP ${res.status}`, text.slice(0, 2000));
    let data;
    try { data = text ? JSON.parse(text) : {}; } catch { throw new TcError(`${service}/${operation}: response is not JSON`, text.slice(0, 2000)); }

    // SOA exceptions come back as a body whose .QName names an Exception type
    if (data['.QName'] && /Exception/.test(data['.QName'])) {
      throw new TcError(data.message || data['.QName'], data);
    }
    const sd = data.ServiceData || data.serviceData || (data.modelObjects || data.partialErrors ? data : null);
    if (sd) this._absorb(sd);
    const errs = sd?.partialErrors || [];
    if (errs.length) {
      const msg = errs.flatMap((e) => (e.errorValues || []).map((v) => v.message)).filter(Boolean).join('; ');
      throw new TcError(msg || `${operation} returned partial errors`, errs);
    }
    return data;
  }

  _absorb(sd) {
    for (const [uid, obj] of Object.entries(sd.modelObjects || {})) {
      const prev = this.objects.get(uid);
      this.objects.set(uid, { ...prev, ...obj, uid, props: { ...(prev?.props || {}), ...(obj.props || {}) } });
    }
  }

  obj(uid) { return this.objects.get(uid); }

  /** Property helpers on cached objects. */
  db(o, name) { const p = (typeof o === 'string' ? this.obj(o) : o)?.props?.[name]; return p?.dbValues || []; }
  ui(o, name) { const p = (typeof o === 'string' ? this.obj(o) : o)?.props?.[name]; return (p?.uiValues || p?.dbValues || [])[0] ?? ''; }

  // ---------------------------------------------------------------- session
  async login(user, password, { group = '', role = '', locale = 'en_US' } = {}) {
    await this.call('Core-2011-06-Session', 'login', {
      credentials: { user, password, group, role, locale, descrimator: 'EPCM-OfficeAddin' }, // (sic) TC spelling
    });
    return this.sessionInfo();
  }

  async sessionInfo() {
    const r = await this.call('Core-2007-01-Session', 'getTCSessionInfo', {});
    this.session = {
      user: r.user ? this.obj(r.user.uid) || r.user : null,
      group: r.group, role: r.role,
      userUid: r.user?.uid,
    };
    if (this.session.userUid) await this.getProperties([r.user], ['user_id', 'user_name', 'object_string']);
    return this.session;
  }

  async logout() {
    try { await this.call('Core-2006-03-Session', 'logout', {}); } finally { this.session = null; this.objects.clear(); }
  }

  // ---------------------------------------------------------------- data
  async getProperties(objects, attributes) {
    if (!objects.length) return [];
    await this.call('Core-2006-03-DataManagement', 'getProperties', { objects: objects.map(ref), attributes });
    return objects.map((o) => this.obj(o.uid) || o);
  }

  /** Run a saved query by name (e.g. 'Item...'), entries/values use the query's display names. */
  async savedQuery(queryName, entries, values, limit = 50) {
    const f = await this.call('Query-2010-04-SavedQuery', 'findSavedQueries', {
      inputCriteria: [{ queryNames: [queryName], queryDescs: [], queryType: 0 }],
    });
    const q = (f.savedQueries || [])[0];
    if (!q) throw new TcError(`Saved query "${queryName}" not found`);
    const r = await this.call('Query-2006-03-SavedQuery', 'executeSavedQuery', {
      query: ref(q), entries, values, limit,
    });
    return (r.objects || []).map((o) => this.obj(o.uid) || o);
  }

  /** Search items by ID / name pattern (wildcards '*'). */
  async searchItems(text, cfg) {
    const pattern = text.includes('*') ? text : `*${text}*`;
    const byId = await this.savedQuery(cfg.itemQuery, [cfg.itemQueryIdEntry], [pattern], cfg.maxResults);
    let byName = [];
    try { byName = await this.savedQuery(cfg.itemQuery, [cfg.itemQueryNameEntry], [pattern], cfg.maxResults); } catch { /* name entry optional */ }
    const seen = new Set(); const items = [];
    for (const o of [...byId, ...byName]) if (!seen.has(o.uid)) { seen.add(o.uid); items.push(o); }
    await this.getProperties(items, ['item_id', 'object_name', 'object_type', 'revision_list', 'object_string']);
    return items.map((o) => this.obj(o.uid));
  }

  async revisions(item) {
    const uids = this.db(item.uid, 'revision_list');
    const revs = uids.map((uid) => ({ uid, type: 'ItemRevision' }));
    await this.getProperties(revs, ['item_revision_id', 'object_string', 'object_name', 'release_status_list']);
    return revs.map((r) => this.obj(r.uid));
  }

  /** Datasets of the given types attached to a revision via the configured relations. */
  async datasets(rev, relations, datasetTypes) {
    await this.getProperties([rev], relations);
    const uids = [...new Set(relations.flatMap((r) => this.db(rev.uid, r)))];
    const ds = uids.map((uid) => ({ uid, type: 'Dataset' }));
    await this.getProperties(ds, DATASET_PROPS);
    return ds.map((d) => this.obj(d.uid)).filter((d) => !datasetTypes || datasetTypes.includes(d.type || this.ui(d, 'object_type')));
  }

  async refreshDataset(uid) {
    await this.getProperties([{ uid, type: 'Dataset' }], DATASET_PROPS);
    return this.obj(uid);
  }

  /** Checkout state relative to the logged-in user. */
  checkoutState(ds) {
    const out = (this.db(ds, 'checked_out')[0] || '').trim().toUpperCase() === 'Y';
    const byUid = this.db(ds, 'checked_out_user')[0];
    return {
      checkedOut: out,
      byMe: out && !!byUid && byUid === this.session?.userUid,
      by: this.ui(ds, 'checked_out_user'),
    };
  }

  // ---------------------------------------------------------------- reservation
  async checkout(objs, comment = 'Checked out from Office') {
    await this.call('Core-2006-03-Reservation', 'checkout', { objects: objs.map(ref), comment, changeId: '' });
  }
  async checkin(objs) {
    await this.call('Core-2006-03-Reservation', 'checkin', { objects: objs.map(ref) });
  }
  async cancelCheckout(objs) {
    await this.call('Core-2006-03-Reservation', 'cancelCheckout', { objects: objs.map(ref) });
  }

  // ---------------------------------------------------------------- files
  /** Download the first named-reference file of a dataset. Returns { bytes: Uint8Array, fileName }. */
  async downloadDataset(ds) {
    await this.getProperties([ds], ['ref_list']);
    const files = this.db(ds.uid, 'ref_list').map((uid) => ({ uid, type: 'ImanFile' }));
    if (!files.length) throw new TcError('Dataset has no file attached');
    await this.getProperties(files, ['original_file_name', 'file_size']);
    const file = files[0];
    const fileName = this.ui(file.uid, 'original_file_name') || 'document';
    const r = await this.call('Core-2006-03-FileManagement', 'getFileReadTickets', { files: [ref(file)] });
    const ticket = mapEntries(r.tickets).map(([, t]) => t)[0];
    if (!ticket) throw new TcError('No read ticket returned by Teamcenter');
    const res = await this._fetch(`${this.baseUrl}${this.fmsPath}/fmsdownload/${encodeURIComponent(fileName)}?ticket=${encodeURIComponent(ticket)}`, { credentials: 'same-origin' });
    if (!res.ok) throw new TcError(`FMS download failed: HTTP ${res.status}`);
    return { bytes: new Uint8Array(await res.arrayBuffer()), fileName };
  }

  /**
   * Upload bytes as the dataset's named reference (replaces the file, new dataset version).
   * Standard thick-client flow: getDatasetWriteTickets -> FMS upload -> commitDatasetFiles.
   */
  async uploadToDataset(ds, bytes, fileName, namedRef) {
    const w = await this.call('Core-2006-03-FileManagement', 'getDatasetWriteTickets', {
      inputs: [{
        dataset: ref(ds), createNewVersion: true,
        datasetFileInfos: [{ clientId: 'f0', fileName, namedReferencedName: namedRef, isText: false, allowReplace: true }],
      }],
    });
    const commit = (w.commitInfo || [])[0];
    const ticket = commit?.datasetFileTicketInfos?.[0]?.ticket;
    if (!ticket) throw new TcError('No write ticket returned by Teamcenter', w);

    const fd = new FormData();
    fd.append('fmsFile', new Blob([bytes], { type: 'application/octet-stream' }), fileName);
    fd.append('fmsTicket', ticket);
    const up = await this._fetch(`${this.baseUrl}${this.fmsPath}/fmsupload/`, { method: 'POST', body: fd, credentials: 'same-origin' });
    if (!up.ok) throw new TcError(`FMS upload failed: HTTP ${up.status}`, await up.text());

    await this.call('Core-2006-03-FileManagement', 'commitDatasetFiles', { commitInput: [commit] });
  }

  // ---------------------------------------------------------------- create
  async newItemId(itemType) {
    const r = await this.call('Core-2006-03-DataManagement', 'generateItemIdsAndInitialRevisionIds', {
      input: [{ itemType, count: 1 }],
    });
    const first = mapEntries(r.outputItemIdsAndInitialRevisionIds)[0]?.[1];
    const v = Array.isArray(first) ? first[0] : first;
    if (!v?.newItemId) throw new TcError('Teamcenter did not generate an item ID');
    return { itemId: v.newItemId, revId: v.newRevId };
  }

  /** Create an Item (+ first revision). Returns { item, rev }. */
  async createItem(itemType, name, description = '') {
    const { itemId, revId } = await this.newItemId(itemType);
    const r = await this.call('Core-2006-03-DataManagement', 'createItems', {
      properties: [{ clientId: 'i0', itemId, itemType, revId, name, description, uom: '', extendedAttributes: [] }],
      container: ref(null), relationType: '',
    });
    const out = (r.output || [])[0];
    if (!out?.item) throw new TcError('Item creation returned no item', r);
    return { item: this.obj(out.item.uid) || out.item, rev: this.obj(out.itemRev.uid) || out.itemRev, itemId, revId };
  }

  /** Create an empty dataset under a revision. */
  async createDataset(rev, datasetType, name, relationType, toolUsed = '') {
    const r = await this.call('Core-2008-06-DataManagement', 'createDatasets2', {
      input: [{ clientId: 'd0', type: datasetType, name, description: '', toolUsed, datasetId: '', datasetRev: '', container: ref(rev), relationType }],
    });
    const ds = (r.output || [])[0]?.dataset;
    if (!ds) throw new TcError('Dataset creation returned no dataset', r);
    await this.refreshDataset(ds.uid);
    return this.obj(ds.uid);
  }
}

export const DATASET_PROPS = ['object_name', 'object_type', 'object_string', 'checked_out', 'checked_out_user', 'last_mod_date', 'ref_list'];
