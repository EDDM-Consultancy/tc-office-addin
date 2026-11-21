// mock-tc.js — in-memory stand-in for the Teamcenter web tier + FSC.
// Implements just the SOA operations and FMS endpoints the add-in uses, with the same
// JSON shapes, so the add-in can be demoed/tested without a Teamcenter server.
//
//   node mock/mock-tc.js            (listens on :7001; login erik / test)
//
// Then run the add-in server with TC_WEB_URL=http://localhost:7001 FSC_URL=http://localhost:7001

const http = require('http');
const crypto = require('crypto');
const JSZip = require('jszip');

const uid = () => crypto.randomBytes(7).toString('base64').replace(/[^A-Za-z0-9]/g, 'x').slice(0, 14);

async function sampleDocx(text) {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`);
  return Buffer.from(await z.generateAsync({ type: 'uint8array' }));
}

function createMock() {
  const db = new Map();        // uid -> { uid, type, className, props: {name: value|array} }
  const files = new Map();     // ImanFile uid -> Buffer
  const tickets = new Map();   // ticket -> { kind: 'read'|'write', file?, data? }
  const sessions = new Map();  // cookie -> userUid
  let nextId = 124;

  const put = (type, props, className = type) => { const o = { uid: uid(), type, className, props }; db.set(o.uid, o); return o; };
  const user = put('User', { user_id: 'erik', user_name: 'Erik', object_string: 'Erik (erik)' });
  const query = put('ImanQuery', { query_name: 'Item...' });

  function addFile(name, buf) { const f = put('ImanFile', { original_file_name: name, file_size: String(buf.length) }); files.set(f.uid, buf); return f; }
  function addDataset(rev, type, name, buf, ext) {
    const d = put(type, { object_name: name, object_type: type, object_string: name, checked_out: '', checked_out_user: null, last_mod_date: new Date().toISOString(), ref_list: [] }, 'Dataset');
    if (buf) d.props.ref_list = [addFile(name + ext, buf).uid];
    rev.props.IMAN_specification = [...(rev.props.IMAN_specification || []), d.uid];
    return d;
  }
  function addItem(itemType, itemId, name, revId = 'A') {
    const item = put(itemType, { item_id: itemId, object_name: name, object_type: itemType, object_string: `${itemId}-${name}`, revision_list: [] }, 'Item');
    const rev = put(itemType + 'Revision', { item_revision_id: revId, object_name: name, object_string: `${itemId}/${revId};1-${name}`, release_status_list: [], IMAN_specification: [], TC_Attaches: [], IMAN_reference: [] }, 'ItemRevision');
    item.props.revision_list.push(rev.uid);
    return { item, rev };
  }

  const ready = (async () => {
    const { rev } = addItem('Document', '000123', 'Pump datasheet');
    addDataset(rev, 'MSWordX', 'Pump datasheet', await sampleDocx('Pump datasheet from mock Teamcenter'), '.docx');
    addItem('Document', '000122', 'Cutter drive calculation');
  })();

  // --- serialisation helpers (TC JSON shapes)
  const toUi = (v) => { if (v == null) return ''; const o = db.get(v); return o ? (o.props.object_string || o.props.user_name || o.uid) : String(v); };
  function mo(o, attrs) {
    const props = {};
    for (const a of attrs || []) {
      if (!(a in o.props)) continue;
      const v = o.props[a];
      const arr = Array.isArray(v) ? v : (v == null ? [] : [v]);
      props[a] = { dbValues: arr.map((x) => (x == null ? '' : String(x))), uiValues: arr.map(toUi) };
    }
    return { objectID: o.uid, uid: o.uid, className: o.className, type: o.type, props };
  }
  const sd = (objs = [], attrs = []) => ({ '.QName': 'http://teamcenter.com/Schemas/Soa/2006-03/Base.ServiceData', plain: objs.map((o) => o.uid), modelObjects: Object.fromEntries(objs.map((o) => [o.uid, mo(o, attrs)])) });
  const r = (o) => ({ uid: o.uid, type: o.type });
  const partial = (message) => ({ ServiceData: { partialErrors: [{ errorValues: [{ message, code: 515000, level: 3 }] }] } });
  const wild = (pat, s) => new RegExp('^' + pat.split('*').map((x) => x.replace(/[.+?^${}()|[\]\\]/g, '\\$&')).join('.*') + '$', 'i').test(s || '');

  const ops = {
    'Core-2011-06-Session/login': (b, ctx) => {
      if (b.credentials.user !== 'erik' || b.credentials.password !== 'test') {
        return { '.QName': 'http://teamcenter.com/Schemas/Soa/2006-03/Exceptions.InvalidCredentialsException', code: 515143, level: 3, message: 'The login attempt failed: either the user ID or the password is invalid.' };
      }
      const sid = crypto.randomBytes(12).toString('hex');
      sessions.set(sid, user.uid);
      ctx.setCookie = `JSESSIONID=${sid}; Path=/tc; HttpOnly`;
      return { serverInfo: { Version: 'P-2412 (mock)' }, ServiceData: sd() };
    },
    'Core-2006-03-Session/logout': (b, ctx) => { sessions.delete(ctx.sid); return { ServiceData: sd() }; },
    'Core-2007-01-Session/getTCSessionInfo': () => ({ user: r(user), group: { uid: 'G', type: 'Group' }, role: { uid: 'R', type: 'Role' }, ServiceData: sd([user], ['user_id', 'user_name']) }),
    'Core-2006-03-DataManagement/getProperties': (b) => {
      const objs = b.objects.map((x) => db.get(x.uid)).filter(Boolean);
      return sd(objs, b.attributes);
    },
    'Query-2010-04-SavedQuery/findSavedQueries': (b) => ({ savedQueries: b.inputCriteria[0].queryNames.includes('Item...') ? [r(query)] : [], ServiceData: sd() }),
    'Query-2006-03-SavedQuery/executeSavedQuery': (b) => {
      const crit = Object.fromEntries(b.entries.map((e, i) => [e, b.values[i]]));
      const hits = [...db.values()].filter((o) => o.className === 'Item'
        && (!crit['Item ID'] || wild(crit['Item ID'], o.props.item_id))
        && (!crit.Name || wild(crit.Name, o.props.object_name))).slice(0, b.limit);
      return { nFound: hits.length, objects: hits.map(r), ServiceData: sd(hits, []) };
    },
    'Core-2006-03-Reservation/checkout': (b, ctx) => {
      for (const x of b.objects) {
        const o = db.get(x.uid);
        if (o.props.checked_out === 'Y') return partial(`${o.props.object_name} is already checked out.`);
        o.props.checked_out = 'Y'; o.props.checked_out_user = ctx.user;
      }
      return { ServiceData: sd() };
    },
    'Core-2006-03-Reservation/checkin': (b, ctx) => {
      for (const x of b.objects) {
        const o = db.get(x.uid);
        if (o.props.checked_out_user !== ctx.user) return partial(`${o.props.object_name} is not checked out by you.`);
        o.props.checked_out = ''; o.props.checked_out_user = null;
      }
      return { ServiceData: sd() };
    },
    'Core-2006-03-Reservation/cancelCheckout': (b, ctx) => ops['Core-2006-03-Reservation/checkin'](b, ctx),
    'Core-2006-03-FileManagement/getFileReadTickets': (b) => {
      const keys = []; const vals = [];
      for (const f of b.files) { const t = 'R' + crypto.randomBytes(16).toString('hex'); tickets.set(t, { kind: 'read', file: f.uid }); keys.push(r(db.get(f.uid))); vals.push(t); }
      return { tickets: [keys, vals], ServiceData: sd() };
    },
    'Core-2006-03-FileManagement/getDatasetWriteTickets': (b, ctx) => {
      const commitInfo = b.inputs.map((inp) => {
        const ds = db.get(inp.dataset.uid);
        if (ds.props.checked_out === 'Y' && ds.props.checked_out_user !== ctx.user) throw Object.assign(new Error('checked out by another user'), { partial: true });
        return {
          dataset: inp.dataset, createNewVersion: inp.createNewVersion,
          datasetFileTicketInfos: inp.datasetFileInfos.map((fi) => { const t = 'W' + crypto.randomBytes(16).toString('hex'); tickets.set(t, { kind: 'write' }); return { datasetFileInfo: fi, ticket: t }; }),
        };
      });
      return { commitInfo, ServiceData: sd() };
    },
    'Core-2006-03-FileManagement/commitDatasetFiles': (b) => {
      for (const ci of b.commitInput) {
        const ds = db.get(ci.dataset.uid);
        const refs = [];
        for (const ti of ci.datasetFileTicketInfos) {
          const t = tickets.get(ti.ticket);
          if (!t?.data) return partial('File was not uploaded to FMS for ticket');
          refs.push(addFile(ti.datasetFileInfo.fileName, t.data).uid);
        }
        ds.props.ref_list = refs; ds.props.last_mod_date = new Date().toISOString();
        ds.versions = (ds.versions || 1) + 1;
      }
      return { ServiceData: sd() };
    },
    'Core-2006-03-DataManagement/generateItemIdsAndInitialRevisionIds': (b) => ({
      outputItemIdsAndInitialRevisionIds: [b.input.map((_, i) => i), b.input.map((inp) => Array.from({ length: inp.count }, () => ({ newItemId: String(nextId++).padStart(6, '0'), newRevId: 'A' })))],
      ServiceData: sd(),
    }),
    'Core-2006-03-DataManagement/createItems': (b) => {
      const output = b.properties.map((p) => { const { item, rev } = addItem(p.itemType, p.itemId, p.name, p.revId); return { clientId: p.clientId, item: r(item), itemRev: r(rev) }; });
      const objs = output.flatMap((o) => [db.get(o.item.uid), db.get(o.itemRev.uid)]);
      return { output, ServiceData: { ...sd(objs, ['item_id', 'item_revision_id', 'object_name']), created: objs.map((o) => o.uid) } };
    },
    'Core-2008-06-DataManagement/createDatasets2': (b) => {
      const output = b.input.map((d) => { const ds = addDataset(db.get(d.container.uid), d.type, d.name, null); return { clientId: d.clientId, dataset: r(ds) }; });
      return { output, ServiceData: sd() };
    },
  };

  async function readBody(req) { const chunks = []; for await (const c of req) chunks.push(c); return Buffer.concat(chunks); }

  const server = http.createServer(async (req, res) => {
    await ready;
    const url = new URL(req.url, 'http://x');
    try {
      // ---------------- FSC
      if (url.pathname.startsWith('/fms/fmsdownload/')) {
        const t = tickets.get(url.searchParams.get('ticket'));
        if (!t || t.kind !== 'read') { res.writeHead(403); return res.end('bad ticket'); }
        res.writeHead(200, { 'Content-Type': 'application/octet-stream' });
        return res.end(files.get(t.file));
      }
      if (url.pathname === '/fms/fmsupload/' && req.method === 'POST') {
        const body = await readBody(req);
        const form = await new Response(body, { headers: { 'content-type': req.headers['content-type'] } }).formData();
        const t = tickets.get(form.get('fmsTicket'));
        if (!t || t.kind !== 'write') { res.writeHead(403); return res.end('bad ticket'); }
        t.data = Buffer.from(await form.get('fmsFile').arrayBuffer());
        res.writeHead(200); return res.end('OK');
      }
      // ---------------- SOA
      const m = url.pathname.match(/^\/tc\/JsonRestServices\/([^/]+)\/([^/]+)$/);
      if (!m || req.method !== 'POST') { res.writeHead(404); return res.end('not found'); }
      const key = `${m[1]}/${m[2]}`;
      const op = ops[key];
      if (!op) { res.writeHead(404); return res.end(`mock: ${key} not implemented`); }
      const env = JSON.parse((await readBody(req)).toString('utf8'));
      const sid = (req.headers.cookie || '').match(/JSESSIONID=([0-9a-f]+)/)?.[1];
      const ctx = { sid, user: sessions.get(sid) };
      let out;
      if (key !== 'Core-2011-06-Session/login' && !ctx.user) {
        out = { '.QName': 'http://teamcenter.com/Schemas/Soa/2006-03/Exceptions.InvalidUserException', code: 515114, level: 3, message: 'The user session is not logged in.' };
      } else {
        try { out = op(env.body, ctx); } catch (e) { out = e.partial ? partial(e.message) : (() => { throw e; })(); }
      }
      const headers = { 'Content-Type': 'application/json' };
      if (ctx.setCookie) headers['Set-Cookie'] = ctx.setCookie;
      res.writeHead(200, headers);
      res.end(JSON.stringify(out));
    } catch (e) {
      res.writeHead(500); res.end(String(e.stack || e));
    }
  });
  return { server, db, files, ready };
}

if (require.main === module) {
  const port = Number(process.env.PORT || 7001);
  createMock().server.listen(port, () => console.log(`Mock Teamcenter on http://localhost:${port}  (login erik / test)`));
}

module.exports = { createMock, sampleDocx };
