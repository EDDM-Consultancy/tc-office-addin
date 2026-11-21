/* global Office, JSZip */
// taskpane.js — UI logic for the Teamcenter Office add-in (Word + Excel).

import { TcClient, TcError } from './tc-client.js';
import * as doc from './office-doc.js';
import { setCustomProps } from './docprops.js';

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const tc = new TcClient();
let cfg;            // /config.json
let hostCfg;        // cfg.hosts[Word|Excel]
let link = null;    // TC link of the current document
let pickedRev = null;

// ------------------------------------------------------------------ ui helpers
function msg(text, kind = 'info') {
  const m = $('msg');
  if (!text) { m.className = 'msg hidden'; return; }
  m.className = `msg ${kind}`; m.textContent = text;
}
async function busy(text, fn) {
  $('busy').classList.remove('hidden'); $('busyText').textContent = text; msg('');
  document.body.classList.add('locked');
  try { return await fn(); }
  catch (e) {
    console.error(e);
    const text = e instanceof TcError ? `Teamcenter: ${e.message}` : (e?.message || String(e));
    msg(text, 'err');
    if (e instanceof TcError && /session|login|credential|401/i.test(e.message)) showLogin();
  } finally {
    $('busy').classList.add('hidden');
    document.body.classList.remove('locked');
  }
}
const show = (id, on) => $(id).classList.toggle('hidden', !on);

function showLogin() { show('loginView', true); show('mainView', false); $('who').textContent = ''; }
function showMain() {
  show('loginView', false); show('mainView', true);
  const u = tc.session?.userUid;
  $('who').textContent = u ? `${tc.ui(u, 'user_name') || tc.ui(u, 'user_id')}` : '';
}

// ------------------------------------------------------------------ init
Office.onReady(async () => {
  cfg = await (await fetch('/config.json')).json();
  hostCfg = cfg.hosts[doc.host()];
  if (!hostCfg) { msg(`Host ${doc.host()} is not supported.`, 'err'); return; }
  const sup = doc.supports();
  if (!sup.open || !sup.props) msg(`This Office version lacks required APIs (${doc.host()}API). Update Office.`, 'err');

  $('newItemType').innerHTML = cfg.newItemTypes.map((t) => `<option>${esc(t)}</option>`).join('');
  $('checkinAfter').checked = !!cfg.checkinAfterSaveDefault;
  wire();

  // Reuse an existing TC session (cookie) if there is one.
  try { await tc.sessionInfo(); showMain(); await loadLink(); }
  catch { showLogin(); }
});

function wire() {
  $('loginForm').addEventListener('submit', (e) => { e.preventDefault(); busy('Signing in…', async () => {
    await tc.login($('user').value.trim(), $('password').value, { group: $('group').value.trim(), role: $('role').value.trim() });
    $('password').value = '';
    showMain(); await loadLink();
  }); });
  $('btnLogout').onclick = () => busy('Signing out…', async () => { await tc.logout(); showLogin(); });

  document.querySelectorAll('.tabs button').forEach((b) => b.onclick = () => {
    document.querySelectorAll('.tabs button').forEach((x) => x.classList.toggle('active', x === b));
    document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('hidden', t.id !== `tab-${b.dataset.tab}`));
  });

  document.querySelectorAll('input[name=saveMode]').forEach((r) => r.onchange = () => {
    const mode = document.querySelector('input[name=saveMode]:checked').value;
    show('saveNew', mode === 'new'); show('saveExisting', mode === 'existing');
  });
  $('newItemName').addEventListener('input', () => { $('dsName').value = $('newItemName').value; });

  $('btnSave').onclick = () => busy('Saving to Teamcenter…', saveLinked);
  $('btnCheckout').onclick = () => busy('Checking out…', async () => { await tc.checkout([dsRef()]); await refreshLinked(); msg('Checked out.', 'ok'); });
  $('btnCheckin').onclick = () => busy('Checking in…', async () => { await tc.checkin([dsRef()]); await refreshLinked(); msg('Checked in.', 'ok'); });
  $('btnCancelCo').onclick = () => busy('Cancelling check-out…', async () => {
    await tc.cancelCheckout([dsRef()]); await refreshLinked();
    msg('Check-out cancelled. Changes you did not save to Teamcenter stay only in this window.', 'ok');
  });
  $('btnRefresh').onclick = () => busy('Refreshing…', refreshLinked);
  $('btnUnlink').onclick = () => busy('Unlinking…', async () => {
    await doc.clearLink(); link = null; renderLink();
    msg('Unlinked. Save to Teamcenter will now create a new dataset.', 'info');
  });
  $('btnSaveNew').onclick = () => busy('Creating in Teamcenter…', saveNew);

  $('revSearchForm').addEventListener('submit', (e) => { e.preventDefault(); busy('Searching…', () => search($('revSearch').value, $('revResults'), 'pick')); });
  $('openSearchForm').addEventListener('submit', (e) => { e.preventDefault(); busy('Searching…', () => search($('openSearch').value, $('openResults'), 'open')); });
}

// ------------------------------------------------------------------ linked document
const dsRef = () => ({ uid: link.datasetUid, type: link.datasetType || hostCfg.datasetType });

async function loadLink() {
  link = await doc.getLink().catch(() => null);
  if (link) await refreshLinked(); else renderLink();
}

async function refreshLinked() {
  if (!link) return renderLink();
  const ds = await tc.refreshDataset(link.datasetUid);
  renderLink(ds);
}

function renderLink(ds) {
  show('linked', !!link); show('unlinked', !link);
  if (!link) { if (!$('dsName').value) $('dsName').value = ''; return; }
  $('lnName').textContent = (ds && tc.ui(ds, 'object_name')) || link.name || '(dataset)';
  $('lnItem').textContent = [link.itemId && `Item ${link.itemId}`, link.revId && `rev ${link.revId}`, link.datasetType].filter(Boolean).join(' · ');
  const st = ds ? tc.checkoutState(ds) : { checkedOut: false };
  const s = $('lnState');
  if (!st.checkedOut) { s.textContent = 'Not checked out'; s.className = 'state free'; }
  else if (st.byMe) { s.textContent = 'Checked out by you'; s.className = 'state mine'; }
  else { s.textContent = `Checked out by ${st.by || 'another user'}`; s.className = 'state other'; }
  $('btnCheckout').disabled = st.checkedOut;
  $('btnCheckin').disabled = !st.byMe;
  $('btnCancelCo').disabled = !st.byMe;
  $('btnSave').disabled = st.checkedOut && !st.byMe;
}

async function saveLinked() {
  const ds = await tc.refreshDataset(link.datasetUid);
  const st = tc.checkoutState(ds);
  if (st.checkedOut && !st.byMe) throw new Error(`Cannot save: dataset is checked out by ${st.by}.`);
  let didCheckout = false;
  if (!st.checkedOut && cfg.autoCheckoutOnSave) { await tc.checkout([ds]); didCheckout = true; }

  const bytes = await doc.getFileBytes();
  await tc.uploadToDataset(ds, bytes, fileName(link.name), hostCfg.namedRef);

  if ($('checkinAfter').checked) await tc.checkin([ds]);
  await refreshLinked();
  msg(`Saved to Teamcenter as a new dataset version${didCheckout && !$('checkinAfter').checked ? ' (dataset is now checked out to you)' : ''}${$('checkinAfter').checked ? ' and checked in' : ''}.`, 'ok');
}

function fileName(name) {
  const ext = doc.fileExt();
  const base = (name || 'document').replace(/[\\/:*?"<>|]/g, '_');
  return base.toLowerCase().endsWith(ext) ? base : base + ext;
}

// ------------------------------------------------------------------ save as new
async function saveNew() {
  const mode = document.querySelector('input[name=saveMode]:checked').value;
  const dsName = $('dsName').value.trim() || $('newItemName').value.trim();
  let rev, itemId, revId;

  if (mode === 'new') {
    const name = $('newItemName').value.trim();
    if (!name) throw new Error('Enter an item name.');
    const r = await tc.createItem($('newItemType').value, name, $('newItemDesc').value.trim());
    ({ rev, itemId, revId } = r);
  } else {
    if (!pickedRev) throw new Error('Search and pick an item revision first.');
    rev = pickedRev.rev; itemId = pickedRev.itemId; revId = pickedRev.revId;
  }
  if (!dsName) throw new Error('Enter a dataset name.');

  const ds = await tc.createDataset(rev, hostCfg.datasetType, dsName, cfg.createRelation, hostCfg.toolUsed || '');
  const newLink = { datasetUid: ds.uid, datasetType: hostCfg.datasetType, itemId, revId, name: dsName };

  // Stamp the link into the open document first, so the uploaded file carries it too.
  await doc.setLink(newLink);
  const bytes = await doc.getFileBytes();
  await tc.uploadToDataset(ds, bytes, fileName(dsName), hostCfg.namedRef);

  link = newLink; pickedRev = null; show('revPicked', false);
  await refreshLinked();
  msg(`Saved to Teamcenter: ${itemId}/${revId} – ${dsName}`, 'ok');
}

// ------------------------------------------------------------------ search / browse
async function search(text, target, mode) {
  text = text.trim();
  if (!text) throw new Error('Enter an item ID or name.');
  const items = await tc.searchItems(text, cfg);
  if (!items.length) { target.innerHTML = '<p class="muted">No items found.</p>'; return; }
  target.innerHTML = items.map((it) => `
    <div class="item" data-uid="${esc(it.uid)}">
      <div class="head"><span><b>${esc(tc.ui(it, 'item_id'))}</b> ${esc(tc.ui(it, 'object_name'))}</span><span class="muted">${esc(tc.ui(it, 'object_type') || it.type)}</span></div>
      <div class="body hidden"></div>
    </div>`).join('');
  target.querySelectorAll('.item > .head').forEach((h) => h.onclick = () => {
    const box = h.parentElement; const body = box.querySelector('.body');
    if (!body.classList.contains('hidden')) { body.classList.add('hidden'); return; }
    busy('Loading…', () => expandItem(box.dataset.uid, body, mode));
  });
}

async function expandItem(itemUid, body, mode) {
  const item = tc.obj(itemUid);
  const revs = (await tc.revisions(item)).reverse(); // newest first
  const itemId = tc.ui(item, 'item_id');
  let html = '';
  for (const rev of revs) {
    const revId = tc.ui(rev, 'item_revision_id');
    const released = tc.db(rev, 'release_status_list').length > 0;
    html += `<div class="rev" data-uid="${esc(rev.uid)}"><div class="rhead">Rev ${esc(revId)}${released ? ' <span class="muted">(released)</span>' : ''}`;
    if (mode === 'pick') {
      html += ` <button class="pick" data-rev="${esc(rev.uid)}" data-item="${esc(itemId)}" data-revid="${esc(revId)}">Use this revision</button></div>`;
    } else {
      html += '</div>';
      const dss = await tc.datasets(rev, cfg.relations, [hostCfg.datasetType]);
      if (!dss.length) html += '<div class="ds muted">No ' + esc(hostCfg.datasetType) + ' datasets</div>';
      for (const d of dss) {
        const st = tc.checkoutState(d);
        html += `<div class="ds"><div><div class="dn">${esc(tc.ui(d, 'object_name'))}</div>${st.checkedOut ? `<div class="co">Checked out by ${st.byMe ? 'you' : esc(st.by)}</div>` : ''}</div>
          <div class="acts">
            <button data-act="ro" data-ds="${esc(d.uid)}" data-item="${esc(itemId)}" data-revid="${esc(revId)}">Open</button>
            ${!st.checkedOut ? `<button data-act="co" class="primary" data-ds="${esc(d.uid)}" data-item="${esc(itemId)}" data-revid="${esc(revId)}">Check out &amp; open</button>` : ''}
          </div></div>`;
      }
    }
    html += '</div>';
  }
  body.innerHTML = html || '<p class="muted">No revisions.</p>';
  body.classList.remove('hidden');

  body.querySelectorAll('button.pick').forEach((b) => b.onclick = () => {
    pickedRev = { rev: tc.obj(b.dataset.rev), itemId: b.dataset.item, revId: b.dataset.revid };
    $('revPicked').textContent = `Selected: ${pickedRev.itemId} / ${pickedRev.revId}`;
    show('revPicked', true);
    if (!$('dsName').value) $('dsName').value = tc.ui(tc.obj(itemUid), 'object_name');
  });
  body.querySelectorAll('button[data-act]').forEach((b) => b.onclick = () =>
    busy(b.dataset.act === 'co' ? 'Checking out and opening…' : 'Opening…', () => openDataset(b.dataset.ds, b.dataset.item, b.dataset.revid, b.dataset.act === 'co')));
}

async function openDataset(dsUid, itemId, revId, checkout) {
  const ds = tc.obj(dsUid);
  if (checkout) await tc.checkout([ds]);
  const { bytes } = await tc.downloadDataset(ds);
  const l = { datasetUid: ds.uid, datasetType: ds.type || hostCfg.datasetType, itemId, revId, name: tc.ui(ds, 'object_name') };
  const stamped = await setCustomProps(JSZip, bytes, doc.linkToProps(l));
  await doc.openNew(stamped);
  msg(`Opened ${l.name} in a new window${checkout ? ' (checked out to you)' : (cfg.autoCheckoutOnSave ? ' (not checked out — it will be checked out to you when you save)' : ' (not checked out)')}.`, 'ok');
}
