/*
 * Lector mínimo de archivos .xlsx, sin dependencias externas.
 *
 * Un .xlsx es un ZIP con XML adentro. Este módulo:
 *   1. Lee el índice del ZIP y descomprime solo las partes necesarias
 *      (usa DecompressionStream, nativo en Edge/Chrome/Firefox/Safari actuales).
 *   2. Interpreta workbook, sharedStrings, styles y cada hoja.
 *   3. Devuelve cada hoja como una matriz de valores (texto, número, booleano o Date).
 *
 * Todo ocurre en el navegador: el archivo no se envía a ningún servidor.
 */
(function (global) {
  "use strict";

  const utf8 = new TextDecoder("utf-8");
  const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  // ---------- ZIP ----------

  function findEndOfCentralDirectory(view) {
    const min = Math.max(0, view.byteLength - 65557);
    for (let i = view.byteLength - 22; i >= min; i--) {
      if (view.getUint32(i, true) === 0x06054b50) return i;
    }
    throw new Error("El archivo no es un Excel válido (.xlsx).");
  }

  function readZipIndex(buffer) {
    const view = new DataView(buffer);
    const eocd = findEndOfCentralDirectory(view);
    const total = view.getUint16(eocd + 10, true);
    let p = view.getUint32(eocd + 16, true);
    const entries = new Map();
    for (let i = 0; i < total; i++) {
      if (view.getUint32(p, true) !== 0x02014b50) {
        throw new Error("El archivo Excel está dañado (índice ZIP inválido).");
      }
      const method = view.getUint16(p + 10, true);
      const compressedSize = view.getUint32(p + 20, true);
      const nameLength = view.getUint16(p + 28, true);
      const extraLength = view.getUint16(p + 30, true);
      const commentLength = view.getUint16(p + 32, true);
      const localOffset = view.getUint32(p + 42, true);
      const name = utf8.decode(new Uint8Array(buffer, p + 46, nameLength));
      entries.set(name, { method, compressedSize, localOffset });
      p += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("Este navegador es muy antiguo para leer Excel. Usá Edge o Chrome actualizados.");
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readZipText(buffer, entries, name) {
    const entry = entries.get(name);
    if (!entry) return null;
    const view = new DataView(buffer);
    const p = entry.localOffset;
    if (view.getUint32(p, true) !== 0x04034b50) {
      throw new Error("El archivo Excel está dañado (" + name + ").");
    }
    const start = p + 30 + view.getUint16(p + 26, true) + view.getUint16(p + 28, true);
    const data = new Uint8Array(buffer, start, entry.compressedSize);
    if (entry.method === 0) return utf8.decode(data);
    if (entry.method === 8) return utf8.decode(await inflateRaw(data));
    throw new Error("Formato de compresión no soportado en " + name + ".");
  }

  // ---------- XML helpers ----------

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) {
      throw new Error("No se pudo interpretar el contenido del Excel.");
    }
    return doc;
  }

  function children(el, localName) {
    const out = [];
    for (const c of el.children) if (c.localName === localName) out.push(c);
    return out;
  }

  function firstChild(el, localName) {
    for (const c of el.children) if (c.localName === localName) return c;
    return null;
  }

  function all(doc, localName) {
    return Array.from(doc.getElementsByTagNameNS("*", localName));
  }

  // Texto de un <si> o <is>: concatena <t> directos y de cada <r>, ignora <rPh> (fonética).
  function richText(el) {
    let s = "";
    for (const c of el.children) {
      if (c.localName === "t") s += c.textContent;
      else if (c.localName === "r") {
        const t = firstChild(c, "t");
        if (t) s += t.textContent;
      }
    }
    return s;
  }

  function resolvePath(target) {
    if (target.startsWith("/")) return target.slice(1);
    // Relativo a xl/
    const parts = ("xl/" + target).split("/");
    const out = [];
    for (const part of parts) {
      if (part === "..") out.pop();
      else if (part !== ".") out.push(part);
    }
    return out.join("/");
  }

  function columnIndex(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const code = ref.charCodeAt(i);
      if (code < 65 || code > 90) break;
      n = n * 26 + (code - 64);
    }
    return n - 1;
  }

  // ---------- Fechas ----------

  const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

  function isDateFormatCode(code) {
    if (!code) return false;
    const cleaned = code
      .replace(/"[^"]*"/g, "")
      .replace(/\[[^\]]*\]/g, "")
      .replace(/\\./g, "")
      .toLowerCase();
    if (cleaned === "general") return false;
    return /[dmyhs]/.test(cleaned);
  }

  function serialToDate(serial, date1904) {
    const days = date1904 ? serial + 1462 : serial;
    // 25569 = días entre 1899-12-30 y 1970-01-01
    return new Date(Math.round((days - 25569) * 86400000));
  }

  // ---------- Lectura principal ----------

  async function readWorkbook(buffer) {
    const entries = readZipIndex(buffer);

    const workbookXml = await readZipText(buffer, entries, "xl/workbook.xml");
    if (!workbookXml) throw new Error("El archivo no parece un Excel (.xlsx).");
    const workbook = parseXml(workbookXml);
    const workbookPr = all(workbook, "workbookPr")[0];
    const date1904 = workbookPr ? ["1", "true"].includes(workbookPr.getAttribute("date1904")) : false;

    const relsXml = await readZipText(buffer, entries, "xl/_rels/workbook.xml.rels");
    const relTargets = new Map();
    if (relsXml) {
      for (const rel of all(parseXml(relsXml), "Relationship")) {
        relTargets.set(rel.getAttribute("Id"), resolvePath(rel.getAttribute("Target")));
      }
    }

    const sharedStrings = [];
    const sstXml = await readZipText(buffer, entries, "xl/sharedStrings.xml");
    if (sstXml) {
      for (const si of all(parseXml(sstXml), "si")) sharedStrings.push(richText(si));
    }

    const dateStyles = new Set();
    const stylesXml = await readZipText(buffer, entries, "xl/styles.xml");
    if (stylesXml) {
      const styles = parseXml(stylesXml);
      const customFormats = new Map();
      for (const nf of all(styles, "numFmt")) {
        customFormats.set(Number(nf.getAttribute("numFmtId")), nf.getAttribute("formatCode"));
      }
      const cellXfs = all(styles, "cellXfs")[0];
      if (cellXfs) {
        children(cellXfs, "xf").forEach((xf, index) => {
          const id = Number(xf.getAttribute("numFmtId") || 0);
          if (BUILTIN_DATE_FORMATS.has(id) || isDateFormatCode(customFormats.get(id))) dateStyles.add(index);
        });
      }
    }

    const sheets = [];
    for (const sheetEl of all(workbook, "sheet")) {
      const name = sheetEl.getAttribute("name");
      const relId = sheetEl.getAttributeNS(NS_REL, "id") || sheetEl.getAttribute("r:id");
      const path = relTargets.get(relId);
      const xml = path ? await readZipText(buffer, entries, path) : null;
      sheets.push({
        name,
        state: sheetEl.getAttribute("state") || "visible",
        rows: xml ? readSheetRows(parseXml(xml), sharedStrings, dateStyles, date1904) : [],
      });
    }
    return { sheets };
  }

  function readSheetRows(doc, sharedStrings, dateStyles, date1904) {
    const rows = [];
    let nextRow = 0;
    for (const rowEl of all(doc, "row")) {
      const r = Number(rowEl.getAttribute("r"));
      const rowIndex = r ? r - 1 : nextRow;
      nextRow = rowIndex + 1;
      const row = [];
      let nextCol = 0;
      for (const c of children(rowEl, "c")) {
        const ref = c.getAttribute("r");
        const col = ref ? columnIndex(ref) : nextCol;
        nextCol = col + 1;
        const value = cellValue(c, sharedStrings, dateStyles, date1904);
        if (value !== null && value !== "") row[col] = value;
      }
      rows[rowIndex] = row;
    }
    // Matriz densa: celdas vacías como null
    const width = rows.reduce((w, row) => Math.max(w, row ? row.length : 0), 0);
    const dense = [];
    for (let i = 0; i < rows.length; i++) {
      const src = rows[i] || [];
      const out = new Array(width).fill(null);
      for (let j = 0; j < src.length; j++) if (src[j] !== undefined) out[j] = src[j];
      dense.push(out);
    }
    return dense;
  }

  function cellValue(c, sharedStrings, dateStyles, date1904) {
    const type = c.getAttribute("t");
    const v = firstChild(c, "v");
    const raw = v ? v.textContent : null;
    switch (type) {
      case "s":
        return raw === null ? null : sharedStrings[Number(raw)] ?? null;
      case "inlineStr": {
        const is = firstChild(c, "is");
        return is ? richText(is) : null;
      }
      case "str":
        return raw;
      case "b":
        return raw === "1";
      case "e":
        return raw; // p. ej. "#N/A"
      default: {
        if (raw === null || raw === "") return null;
        const num = Number(raw);
        if (Number.isNaN(num)) return raw;
        const style = Number(c.getAttribute("s") || 0);
        return dateStyles.has(style) ? serialToDate(num, date1904) : num;
      }
    }
  }

  global.XlsxReader = { readWorkbook };
})(window);
