#!/usr/bin/env node
/**
 * docx-edit — in-place text replacement inside a .docx that preserves formatting.
 *
 * A .docx is a zip of XML parts. Visible text lives in <w:t> nodes, but Word
 * frequently splits one visible phrase across several runs (e.g. "Skylar" and
 * " Higgins", or "$" / "2,500"). A naive string replace on document.xml misses
 * those. This script joins the text across runs, replaces, then writes the new
 * text back into the FIRST affected run and blanks the rest — so styling,
 * tables, headers, numbering, and every untouched byte are preserved exactly.
 *
 * Usage:
 *   node scripts/docx-edit.mjs <input.docx> -o <output.docx> -r "Old=>New" [-r ...]
 *   node scripts/docx-edit.mjs <input.docx> -o <output.docx> --map repl.json
 *   node scripts/docx-edit.mjs <input.docx> --in-place -r "2025=>2026"
 *   node scripts/docx-edit.mjs <input.docx> --dry-run -r "Old=>New"
 *
 *   --map repl.json   JSON: [{"find":"..","replace":"..","all":true}]  or  {"old":"new"}
 *   -r "A=>B"         repeatable inline rule (replaces ALL occurrences by default)
 *   --first           only replace the first occurrence of each rule
 *   --dry-run         report match counts, write nothing
 *   --allow-missing   don't exit non-zero when a rule matches 0 times
 *   --parts <glob>    comma-list of zip parts to edit
 *                     (default: word/document.xml,word/header*.xml,word/footer*.xml)
 *
 * Exit codes: 0 ok · 1 a rule matched nothing (unless --allow-missing) · 2 usage/error.
 */

import fs from "node:fs";
import JSZip from "jszip";

// ---- tiny XML text entity codecs (element-content only: & < >) -------------
const decodeXml = (s) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">")
   .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
   .replace(/&amp;/g, "&");
const encodeXml = (s) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// ---- cross-run replacement within a single XML part ------------------------
// Returns { xml, count }. Untouched <w:t> nodes are re-emitted byte-for-byte.
function replaceInPart(xml, rules) {
  const re = /<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/g;
  const nodes = [];
  let m;
  while ((m = re.exec(xml)) !== null) {
    nodes.push({
      start: m.index,
      end: m.index + m[0].length,
      attrs: m[1],
      text: decodeXml(m[2]),
      dirty: false,
    });
  }

  const counts = {};
  for (const rule of rules) {
    const { find, replace, all } = rule;
    if (!find) continue;
    let searchFrom = 0;
    let n = 0;
    for (;;) {
      const full = nodes.map((x) => x.text).join("");
      const idx = full.indexOf(find, searchFrom);
      if (idx === -1) break;
      const endIdx = idx + find.length;

      // map [idx, endIdx) onto node indices/offsets
      let pos = 0, sNode = -1, sOff = 0, eNode = -1, eOff = 0;
      for (let i = 0; i < nodes.length; i++) {
        const len = nodes[i].text.length;
        if (sNode === -1 && idx < pos + len) { sNode = i; sOff = idx - pos; }
        if (endIdx <= pos + len) { eNode = i; eOff = endIdx - pos; break; }
        pos += len;
      }
      if (sNode === -1 || eNode === -1) break; // safety

      const prefix = nodes[sNode].text.slice(0, sOff);
      const suffix = nodes[eNode].text.slice(eOff);
      if (sNode === eNode) {
        nodes[sNode].text = prefix + replace + suffix;
      } else {
        nodes[sNode].text = prefix + replace;
        for (let i = sNode + 1; i < eNode; i++) nodes[i].text = "";
        nodes[eNode].text = suffix;
      }
      for (let i = sNode; i <= eNode; i++) nodes[i].dirty = true;

      // advance past the inserted text so we never re-match inside it
      searchFrom = idx + replace.length;
      n++;
      if (!all) break;
    }
    counts[find] = (counts[find] || 0) + n;
  }

  // rebuild: original bytes for clean nodes, re-encoded text for dirty ones
  let out = "", cursor = 0, total = 0;
  for (const node of nodes) {
    out += xml.slice(cursor, node.start);
    if (node.dirty) {
      let attrs = node.attrs;
      if (/^\s|\s$/.test(node.text) && !/xml:space=/.test(attrs)) {
        attrs += ' xml:space="preserve"';
      }
      out += `<w:t${attrs}>${encodeXml(node.text)}</w:t>`;
      total++;
    } else {
      out += xml.slice(node.start, node.end);
    }
    cursor = node.end;
  }
  out += xml.slice(cursor);
  return { xml: out, counts, dirtyNodes: total };
}

// ---- arg parsing -----------------------------------------------------------
function parseArgs(argv) {
  const a = { rules: [], parts: null, out: null, inPlace: false, dryRun: false, allowMissing: false, first: false, input: null };
  for (let i = 0; i < argv.length; i++) {
    const t = argv[i];
    if (t === "-o" || t === "--out") a.out = argv[++i];
    else if (t === "--in-place") a.inPlace = true;
    else if (t === "--dry-run") a.dryRun = true;
    else if (t === "--allow-missing") a.allowMissing = true;
    else if (t === "--first") a.first = true;
    else if (t === "--parts") a.parts = argv[++i].split(",").map((s) => s.trim()).filter(Boolean);
    else if (t === "-r" || t === "--replace") {
      const spec = argv[++i];
      const sep = spec.indexOf("=>");
      if (sep === -1) { console.error(`Bad -r rule (need "old=>new"): ${spec}`); process.exit(2); }
      a.rules.push({ find: spec.slice(0, sep), replace: spec.slice(sep + 2) });
    } else if (t === "--map") {
      const raw = JSON.parse(fs.readFileSync(argv[++i], "utf-8"));
      if (Array.isArray(raw)) for (const r of raw) a.rules.push({ find: r.find, replace: r.replace ?? "", all: r.all });
      else for (const [k, v] of Object.entries(raw)) a.rules.push({ find: k, replace: String(v) });
    } else if (t === "-h" || t === "--help") { printHelp(); process.exit(0); }
    else if (!t.startsWith("-") && !a.input) a.input = t;
    else { console.error(`Unknown arg: ${t}`); process.exit(2); }
  }
  return a;
}

function printHelp() {
  const header = fs.readFileSync(new URL(import.meta.url), "utf-8")
    .split("\n").filter((l) => l.startsWith(" *") || l.startsWith("/**"))
    .map((l) => l.replace(/^\/?\*+ ?/, "")).join("\n");
  console.log(header);
}

// ---- main ------------------------------------------------------------------
async function main() {
  const a = parseArgs(process.argv.slice(2));
  if (!a.input) { console.error("Missing input .docx. Use --help."); process.exit(2); }
  if (!a.rules.length) { console.error("No replacement rules. Pass -r or --map."); process.exit(2); }
  if (!a.dryRun && !a.out && !a.inPlace) { console.error("Specify -o <output.docx>, --in-place, or --dry-run."); process.exit(2); }
  const outPath = a.inPlace ? a.input : a.out;

  // default "all" unless --first; let per-rule .all (from --map) win when set
  for (const r of a.rules) if (r.all === undefined) r.all = !a.first;

  const zip = await JSZip.loadAsync(fs.readFileSync(a.input));
  const defaults = [/^word\/document\.xml$/, /^word\/header\d*\.xml$/, /^word\/footer\d*\.xml$/];
  const wanted = a.parts
    ? (name) => a.parts.includes(name)
    : (name) => defaults.some((rx) => rx.test(name));

  const tally = {};
  for (const r of a.rules) tally[r.find] = 0;
  let editedParts = 0;

  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir || !wanted(name)) continue;
    const xml = await zip.files[name].async("string");
    const { xml: newXml, counts, dirtyNodes } = replaceInPart(xml, a.rules);
    for (const [k, v] of Object.entries(counts)) tally[k] += v;
    if (dirtyNodes > 0) {
      editedParts++;
      if (!a.dryRun) zip.file(name, newXml);
    }
  }

  // report
  let missing = false;
  console.log(`${a.dryRun ? "[dry-run] " : ""}${a.input}`);
  for (const r of a.rules) {
    const c = tally[r.find];
    if (c === 0) missing = true;
    console.log(`  ${c === 0 ? "✗" : "✓"} ${c}×  "${r.find}" → "${r.replace}"`);
  }
  console.log(`  parts changed: ${editedParts}`);

  if (!a.dryRun && editedParts > 0) {
    const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    fs.writeFileSync(outPath, buf);
    console.log(`  wrote: ${outPath} (${buf.length} bytes)`);
  } else if (!a.dryRun) {
    console.log("  nothing to write (no matches).");
  }

  if (missing && !a.allowMissing) {
    console.error("One or more rules matched nothing — check exact text (Word may use curly quotes/non-breaking spaces). Use --dry-run to probe, or --allow-missing to ignore.");
    process.exit(1);
  }
}

main().catch((e) => { console.error(e?.stack || String(e)); process.exit(2); });
