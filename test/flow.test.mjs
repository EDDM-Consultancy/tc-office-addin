// End-to-end test: TcClient -> add-in proxy server -> mock Teamcenter (SOA + FMS).
// Run: npm test
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { TcClient, TcError } from '../src/tc-client.js';
import { setCustomProps, getCustomProps } from '../src/docprops.js';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const { createMock, sampleDocx } = require('../mock/mock-tc.js');
const { createApp } = require('../server.js');
const cfgFile = require('../config.json');

let mock, proxy, base;
const cfg = cfgFile.client;

function cookieFetch() {
  const jar = new Map();
  return async (url, opts = {}) => {
    const headers = new Headers(opts.headers || {});
    if (jar.size) headers.set('cookie', [...jar].map(([k, v]) => `${k}=${v}`).join('; '));
    const res = await fetch(url, { ...opts, headers });
    for (const c of res.headers.getSetCookie?.() || []) {
      const [kv] = c.split(';'); const [k, v] = kv.split('=');
      jar.set(k.trim(), v);
      assert.match(c, /Path=\/tc/i, 'cookie path must stay /tc through the proxy');
    }
    return res;
  };
}

const listen = (srv) => new Promise((r) => srv.listen(0, () => r(srv.address().port)));

before(async () => {
  mock = createMock(); await mock.ready;
  const mp = await listen(mock.server);
  const app = createApp({ server: { tcWebUrl: `http://localhost:${mp}`, tcContext: 'tc', fscUrl: `http://localhost:${mp}` }, client: cfg });
  proxy = (await import('node:http')).createServer(app);
  base = `http://localhost:${await listen(proxy)}`;
});
after(() => { mock.server.close(); proxy.close(); });

test('wrong password raises TcError', async () => {
  const tc = new TcClient({ baseUrl: base, fetch: cookieFetch() });
  await assert.rejects(tc.login('erik', 'nope'), (e) => e instanceof TcError && /password is invalid/.test(e.message));
});

test('calls without a session are rejected', async () => {
  const tc = new TcClient({ baseUrl: base, fetch: cookieFetch() });
  await assert.rejects(tc.sessionInfo(), /not logged in/);
});

test('serves config and task pane', async () => {
  const c = await (await fetch(`${base}/config.json`)).json();
  assert.equal(c.hosts.Word.datasetType, 'MSWordX');
  const html = await (await fetch(`${base}/`)).text();
  assert.match(html, /taskpane\.js/);
  assert.equal((await fetch(`${base}/vendor/jszip.min.js`)).status, 200);
});

test('open → check out → edit → save → check in', async () => {
  const tc = new TcClient({ baseUrl: base, fetch: cookieFetch() });
  const s = await tc.login('erik', 'test');
  assert.ok(s.userUid);

  const items = await tc.searchItems('Pump', cfg);
  assert.equal(items.length, 1);
  assert.equal(tc.ui(items[0], 'item_id'), '000123');

  const [rev] = await tc.revisions(items[0]);
  const dss = await tc.datasets(rev, cfg.relations, ['MSWordX']);
  assert.equal(dss.length, 1);
  const ds = dss[0];
  assert.equal(tc.checkoutState(ds).checkedOut, false);

  await tc.checkout([ds]);
  assert.deepEqual(tc.checkoutState(await tc.refreshDataset(ds.uid)), { checkedOut: true, byMe: true, by: 'Erik (erik)' });

  const { bytes, fileName } = await tc.downloadDataset(ds);
  assert.equal(fileName, 'Pump datasheet.docx');
  const stamped = await setCustomProps(JSZip, bytes, { TC_DatasetUID: ds.uid, TC_ItemID: '000123' });
  assert.equal((await getCustomProps(JSZip, stamped)).TC_DatasetUID, ds.uid);
  // still a valid package with custom part registered
  const z = await JSZip.loadAsync(stamped);
  assert.match(await z.file('[Content_Types].xml').async('string'), /docProps\/custom\.xml/);
  assert.match(await z.file('_rels/.rels').async('string'), /custom-properties/);

  const edited = await sampleDocx('Edited in Word');
  await tc.uploadToDataset(ds, edited, 'Pump datasheet.docx', 'word');
  const back = await tc.downloadDataset(ds);
  assert.deepEqual(Buffer.from(back.bytes), edited);

  await tc.checkin([ds]);
  assert.equal(tc.checkoutState(await tc.refreshDataset(ds.uid)).checkedOut, false);
  await tc.logout();
});

test('save as new item, and save onto existing revision', async () => {
  const tc = new TcClient({ baseUrl: base, fetch: cookieFetch() });
  await tc.login('erik', 'test');

  const { rev, itemId, revId } = await tc.createItem('Document', 'Winch load case', 'from Excel');
  assert.match(itemId, /^\d{6}$/); assert.equal(revId, 'A');
  const ds = await tc.createDataset(rev, 'MSExcelX', 'Winch load case', 'IMAN_specification');
  const bytes = await setCustomProps(JSZip, await sampleDocx('pretend xlsx'), { TC_DatasetUID: ds.uid });
  await tc.uploadToDataset(ds, bytes, 'Winch load case.xlsx', 'excel');
  const found = await tc.datasets(rev, cfg.relations, ['MSExcelX']);
  assert.equal(found.length, 1);
  assert.equal((await getCustomProps(JSZip, (await tc.downloadDataset(found[0])).bytes)).TC_DatasetUID, ds.uid);

  const [cutter] = await tc.searchItems('000122', cfg);
  const [crev] = await tc.revisions(cutter);
  const ds2 = await tc.createDataset(crev, 'MSWordX', 'Cutter drive note', 'IMAN_specification');
  await tc.uploadToDataset(ds2, await sampleDocx('note'), 'Cutter drive note.docx', 'word');
  assert.equal((await tc.datasets(crev, cfg.relations, ['MSWordX'])).length, 1);
});

test('dataset checked out by someone else cannot be checked out again', async () => {
  const a = new TcClient({ baseUrl: base, fetch: cookieFetch() });
  await a.login('erik', 'test');
  const [it] = await a.searchItems('000123', cfg);
  const [rev] = await a.revisions(it);
  const [ds] = await a.datasets(rev, cfg.relations, ['MSWordX']);
  await a.checkout([ds]);
  await assert.rejects(a.checkout([ds]), /already checked out/);
  await a.cancelCheckout([ds]);
});
