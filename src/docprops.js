// docprops.js — read/write custom document properties inside an OOXML package (.docx/.xlsx).
//
// Used to stamp a file downloaded from Teamcenter with the dataset UID before it is opened
// in a new Word/Excel window, so the add-in in that window knows where to save it back.
// JSZip is passed in (browser: window.JSZip, Node tests: require('jszip')).

const CUSTOM_PART = 'docProps/custom.xml';
const CUSTOM_CT = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const CUSTOM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Parse <property name=".."><vt:lpwstr>..</vt:lpwstr></property> entries (string props only). */
export function parseCustomXml(xml) {
  const out = {};
  const re = /<property\b[^>]*\bname="([^"]*)"[^>]*>\s*<vt:(\w+)>([\s\S]*?)<\/vt:\2>\s*<\/property>/g;
  let m;
  while ((m = re.exec(xml))) out[m[1]] = m[3].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
  return out;
}

function buildCustomXml(props) {
  let pid = 2; // pids start at 2
  const body = Object.entries(props).map(([k, v]) =>
    `<property fmtid="${FMTID}" pid="${pid++}" name="${esc(k)}"><vt:lpwstr>${esc(v)}</vt:lpwstr></property>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${body}</Properties>`;
}

/** Return a new package (Uint8Array) with the given string custom properties set/overwritten. */
export async function setCustomProps(JSZip, bytes, props) {
  const zip = await JSZip.loadAsync(bytes);
  const existing = zip.file(CUSTOM_PART) ? parseCustomXml(await zip.file(CUSTOM_PART).async('string')) : {};
  zip.file(CUSTOM_PART, buildCustomXml({ ...existing, ...props }));

  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!ct.includes('/docProps/custom.xml')) {
    ct = ct.replace('</Types>', `<Override PartName="/docProps/custom.xml" ContentType="${CUSTOM_CT}"/></Types>`);
    zip.file('[Content_Types].xml', ct);
  }
  let rels = await zip.file('_rels/.rels').async('string');
  if (!rels.includes(CUSTOM_REL)) {
    rels = rels.replace('</Relationships>', `<Relationship Id="rIdTcCustom" Type="${CUSTOM_REL}" Target="docProps/custom.xml"/></Relationships>`);
    zip.file('_rels/.rels', rels);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

export async function getCustomProps(JSZip, bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const f = zip.file(CUSTOM_PART);
  return f ? parseCustomXml(await f.async('string')) : {};
}
