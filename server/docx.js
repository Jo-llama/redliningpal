// .docx helpers for the server: paragraphs (before/after tracked changes)
// and the rc-ver stamp in docProps/custom.xml.

import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { parseParagraphs, norm } from "../public/lib/ooxml.js";

export const STAMP_NAME = "rc-ver";

export async function readDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const docXml = zip.file("word/document.xml");
  if (!docXml) throw new Error("not a Word document (word/document.xml missing)");
  const paragraphs = parseParagraphs(await docXml.async("string"), DOMParser);
  const custom = zip.file("docProps/custom.xml");
  const stamp = custom ? readStampXml(await custom.async("string")) : null;
  return {
    paragraphs,
    stamp,
    hasTrackedChanges: paragraphs.some((p) => p.hasChanges),
    beforeTexts: paragraphs.map((p) => norm(p.before)), // "reject all" state
    afterTexts: paragraphs.map((p) => norm(p.after)),   // "accept all" state
  };
}

function readStampXml(xml) {
  const m = xml.match(new RegExp(`name="${STAMP_NAME}"[^>]*>\\s*<vt:lpwstr>([^<]+)</vt:lpwstr>`));
  return m ? m[1].trim() : null;
}

/** Return a copy of the .docx with rc-ver set to `value` (add or replace). */
export async function writeStamp(buf, value) {
  const zip = await JSZip.loadAsync(buf);
  const prop = `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="__PID__" name="${STAMP_NAME}"><vt:lpwstr>${value}</vt:lpwstr></property>`;
  const existing = zip.file("docProps/custom.xml");
  let xml;
  if (existing) {
    xml = await existing.async("string");
    const re = new RegExp(`<property[^>]*name="${STAMP_NAME}"[^>]*>[\\s\\S]*?</property>`);
    if (re.test(xml)) {
      xml = xml.replace(re, (m) => m.replace(/<vt:lpwstr>[^<]*<\/vt:lpwstr>/, `<vt:lpwstr>${value}</vt:lpwstr>`));
    } else {
      const pids = [...xml.matchAll(/pid="(\d+)"/g)].map((m) => Number(m[1]));
      xml = xml.replace("</Properties>", prop.replace("__PID__", String(Math.max(1, ...pids) + 1)) + "</Properties>");
    }
  } else {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      prop.replace("__PID__", "2") + "</Properties>";
    let ct = await zip.file("[Content_Types].xml").async("string");
    if (!ct.includes("/docProps/custom.xml")) {
      ct = ct.replace("</Types>", '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
      zip.file("[Content_Types].xml", ct);
    }
    let rels = await zip.file("_rels/.rels").async("string");
    if (!rels.includes("custom-properties")) {
      rels = rels.replace("</Relationships>", '<Relationship Id="rIdRcCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>');
      zip.file("_rels/.rels", rels);
    }
  }
  zip.file("docProps/custom.xml", xml);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
