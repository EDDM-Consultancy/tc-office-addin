/* global Office, Word, Excel */
// office-doc.js — thin abstraction over Word and Excel so the task pane code is host-agnostic.

export const LINK_KEYS = {
  datasetUid: 'TC_DatasetUID',
  datasetType: 'TC_DatasetType',
  itemId: 'TC_ItemID',
  revId: 'TC_RevID',
  name: 'TC_DatasetName',
};

export function host() {
  const h = Office.context.host;
  if (h === Office.HostType.Word) return 'Word';
  if (h === Office.HostType.Excel) return 'Excel';
  return String(h);
}

export function supports() {
  const h = host();
  return h === 'Word'
    ? { open: Office.context.requirements.isSetSupported('WordApi', '1.3'), props: Office.context.requirements.isSetSupported('WordApi', '1.3') }
    : { open: Office.context.requirements.isSetSupported('ExcelApi', '1.8'), props: Office.context.requirements.isSetSupported('ExcelApi', '1.12') };
}

// ------------------------------------------------------------ base64
export function toBase64(bytes) {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(s);
}

// ------------------------------------------------------------ custom properties (TC link)
async function readCustom() {
  if (host() === 'Word') {
    return Word.run(async (ctx) => {
      const c = ctx.document.properties.customProperties;
      c.load('items/key,items/value');
      await ctx.sync();
      return Object.fromEntries(c.items.map((p) => [p.key, String(p.value)]));
    });
  }
  return Excel.run(async (ctx) => {
    const c = ctx.workbook.properties.custom;
    c.load('items/key,items/value');
    await ctx.sync();
    return Object.fromEntries(c.items.map((p) => [p.key, String(p.value)]));
  });
}

async function writeCustom(props) {
  const entries = Object.entries(props);
  if (host() === 'Word') {
    return Word.run(async (ctx) => {
      const c = ctx.document.properties.customProperties;
      for (const [k, v] of entries) c.add(k, v ?? '');
      await ctx.sync();
    });
  }
  return Excel.run(async (ctx) => {
    const c = ctx.workbook.properties.custom;
    for (const [k, v] of entries) c.add(k, v ?? '');
    await ctx.sync();
  });
}

/** The Teamcenter dataset this document is linked to, or null. */
export async function getLink() {
  const p = await readCustom();
  if (!p[LINK_KEYS.datasetUid]) return null;
  return Object.fromEntries(Object.entries(LINK_KEYS).map(([k, key]) => [k, p[key] || '']));
}

export async function setLink(link) {
  await writeCustom(Object.fromEntries(Object.entries(LINK_KEYS).map(([k, key]) => [key, link[k] || ''])));
}

/** Remove the TC link properties from the open document. */
export async function clearLink() {
  const keys = Object.values(LINK_KEYS);
  const run = host() === 'Word' ? Word.run : Excel.run;
  return run(async (ctx) => {
    const c = host() === 'Word' ? ctx.document.properties.customProperties : ctx.workbook.properties.custom;
    const items = keys.map((k) => c.getItemOrNullObject(k));
    items.forEach((i) => i.load('isNullObject'));
    await ctx.sync();
    items.forEach((i) => { if (!i.isNullObject) i.delete(); });
    await ctx.sync();
  });
}

export function linkToProps(link) {
  return Object.fromEntries(Object.entries(LINK_KEYS).map(([k, key]) => [key, link[k] || '']));
}

// ------------------------------------------------------------ file content
/** Current document (including unsaved edits) as .docx/.xlsx bytes. */
export function getFileBytes(sliceSize = 4 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      const file = r.value;
      const parts = new Array(file.sliceCount);
      let got = 0;
      const done = (err) => { file.closeAsync(); err ? reject(err) : resolve(concat(parts)); };
      for (let i = 0; i < file.sliceCount; i++) {
        file.getSliceAsync(i, (s) => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) return done(s.error);
          parts[s.value.index] = new Uint8Array(s.value.data);
          if (++got === file.sliceCount) done();
        });
      }
    });
  });
}

function concat(parts) {
  const len = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

/** Open bytes as a new document window. */
export async function openNew(bytes) {
  const b64 = toBase64(bytes);
  if (host() === 'Word') {
    return Word.run(async (ctx) => { ctx.application.createDocument(b64).open(); await ctx.sync(); });
  }
  return Excel.createWorkbook(b64);
}

export const fileExt = () => (host() === 'Word' ? '.docx' : '.xlsx');
