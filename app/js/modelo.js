/*
 * Convierte el Excel DataGob_Reglas_Negocio.xlsx en el modelo que muestra el tablero.
 *
 * Las columnas se buscan por NOMBRE de encabezado (no por posición), así que el
 * orden de las columnas en el Excel puede cambiar sin romper el tablero.
 * Si falta una hoja o una columna obligatoria, se informa con un mensaje claro.
 */
(function (global) {
  "use strict";

  // ---------- Utilidades ----------

  function norm(s) {
    return String(s ?? "")
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  function clean(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === "string") {
      const s = v.replace(/\s+/g, " ").trim();
      return s === "" ? null : s;
    }
    return v;
  }

  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  }

  function wordCount(s) {
    return String(s).trim().split(/\s+/).filter(Boolean).length;
  }

  function findSheet(workbook, wanted) {
    const target = norm(wanted);
    return (
      workbook.sheets.find((s) => norm(s.name) === target) ||
      workbook.sheets.find((s) => norm(s.name).startsWith(target)) ||
      null
    );
  }

  function requireSheet(workbook, wanted, problems) {
    const sheet = findSheet(workbook, wanted);
    if (!sheet) problems.push(`No se encontró la hoja "${wanted}" en el Excel.`);
    return sheet;
  }

  /*
   * Lee una hoja con formato "tabla": la fila 1 tiene los encabezados y la tabla
   * termina en la primera columna sin encabezado. Lo que está a la derecha, o en
   * filas con una sola celda, se toma como nota explicativa del Excel.
   */
  function readTable(sheet) {
    const rows = sheet ? sheet.rows : [];
    const header = rows[0] || [];
    const columns = [];
    for (let i = 0; i < header.length; i++) {
      if (isEmpty(header[i])) break;
      columns.push({ index: i, key: norm(header[i]) });
    }
    const width = columns.length;
    const records = [];
    const notes = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const filled = columns.filter((c) => !isEmpty(row[c.index])).length;
      if (filled >= 2) {
        const rec = { _fila: r + 1 };
        for (const c of columns) if (!(c.key in rec)) rec[c.key] = clean(row[c.index]);
        records.push(rec);
      } else if (filled === 1) {
        const only = columns.find((c) => !isEmpty(row[c.index]));
        const value = row[only.index];
        if (typeof value === "string" && wordCount(value) >= 3) notes.push(value.trim());
      }
      for (let i = width; i < row.length; i++) {
        const value = row[i];
        if (typeof value === "string" && wordCount(value) >= 3) notes.push(value.trim());
      }
    }
    return { keys: new Set(columns.map((c) => c.key)), records, notes: [...new Set(notes)] };
  }

  function checkColumns(table, sheetName, required, problems) {
    if (!table.records.length && !table.keys.size) return;
    const missing = required.filter((k) => !table.keys.has(norm(k)));
    if (missing.length) {
      problems.push(`En la hoja "${sheetName}" faltan las columnas: ${missing.join(", ")}.`);
    }
  }

  function get(rec, key) {
    return rec[norm(key)] ?? null;
  }

  function isActive(rec) {
    return norm(get(rec, "ACTIVA")) === "SI";
  }

  // "A DEFINIR", vacío o "-" cuentan como dato pendiente
  function isDefined(v) {
    if (isEmpty(v)) return false;
    const s = norm(v);
    return s !== "-" && !s.includes("DEFINIR");
  }

  function code(v) {
    if (isEmpty(v)) return null;
    const s = String(v).trim();
    return s === "-" ? null : s;
  }

  // ---------- Hojas ----------

  const SEDE_CAMPOS = [
    { key: "NOMBRE_SEDE_PARSEADO", label: "Sede normalizada" },
    { key: "TIPO_SEDE", label: "Tipo de sede" },
    { key: "EMPRESA", label: "Empresa" },
    { key: "MODALIDAD", label: "Modalidad" },
    { key: "MODELO_NEGOCIO", label: "Modelo de negocio" },
  ];

  function buildSedes(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["NOMBRE_SEDE", "FUENTE", ...SEDE_CAMPOS.map((c) => c.key)], problems);
    const items = t.records.map((r) => {
      const campos = {};
      const faltan = [];
      for (const c of SEDE_CAMPOS) {
        const v = get(r, c.key);
        campos[c.key] = v;
        if (!isDefined(v)) faltan.push(c.label);
      }
      return {
        numero: get(r, "NUMERO"),
        origenId: get(r, "ORIGEN_ID"),
        nombre: get(r, "NOMBRE_SEDE"),
        fuente: get(r, "FUENTE") || "Sin fuente",
        codigo: get(r, "CODIGO_ORIGEN"),
        observaciones: get(r, "OBSERVACIONES_GOBIERNO"),
        campos,
        faltan,
        completa: faltan.length === 0,
      };
    });
    return { items, notes: t.notes };
  }

  function buildEspecialidadServicio(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["ACTIVA", "FUENTE", "SERVICIO_FUENTE", "ESPECIALIDAD_FUENTE", "SERVICIO_NUEVO", "ESPECIALIDAD_NUEVA"], problems);
    const items = t.records.map((r) => ({
      activa: isActive(r),
      fuente: get(r, "FUENTE"),
      servicio: get(r, "SERVICIO_FUENTE"),
      servicioId: get(r, "ID_SERVICIO_FUENTE"),
      especialidad: get(r, "ESPECIALIDAD_FUENTE"),
      especialidadId: get(r, "ID_ESPECIALIDAD_FUENTE"),
      servicioNuevo: get(r, "SERVICIO_NUEVO"),
      servicioNuevoId: get(r, "ID_SERVICIO_NUEVO"),
      especialidadNueva: get(r, "ESPECIALIDAD_NUEVA"),
      especialidadNuevaId: get(r, "ID_ESPECIALIDAD_NUEVA"),
    }));
    return { items, notes: t.notes };
  }

  function buildUnidadNegocio(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["ACTIVA", "FUENTE", "UNIDAD_NEGOCIO_ACTUAL", "UNIDAD_NEGOCIO_NUEVA"], problems);
    const items = t.records.map((r) => ({
      activa: isActive(r),
      fuente: get(r, "FUENTE"),
      unidadActual: get(r, "UNIDAD_NEGOCIO_ACTUAL"),
      unidadNueva: get(r, "UNIDAD_NEGOCIO_NUEVA"),
      servicio: get(r, "SERVICIO ACTUAL"),
      servicioId: get(r, "ID_SERVICIO_FUENTE"),
      especialidad: get(r, "ESPECIALIDAD ACTUAL"),
      especialidadId: get(r, "ID_ESPECIALIDAD_FUENTE"),
    }));
    return { items, notes: t.notes };
  }

  function buildEstudios(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["ACTIVA", "FUENTE", "PRA_DESCRIPCION", "ESPECIALIDAD_NUEVA", "SERVICIO_NUEVO"], problems);
    const items = t.records.map((r) => ({
      activa: isActive(r),
      fuente: get(r, "FUENTE"),
      praId: get(r, "PRA_ID"),
      praCodigo: get(r, "PRA_CODIGO"),
      praRegion: get(r, "PRA_REGION"),
      estudio: get(r, "PRA_DESCRIPCION"),
      especialidad: get(r, "ESPECIALIDAD_FUENTE"),
      especialidadId: get(r, "ESPECIALIDAD_ID_FUENTE"),
      especialidadNueva: get(r, "ESPECIALIDAD_NUEVA"),
      especialidadNuevaId: get(r, "ID_ESPECIALIDAD_NUEVA"),
      servicioNuevo: get(r, "SERVICIO_NUEVO"),
      servicioNuevoId: get(r, "ID_SERVICIO_NUEVO"),
    }));
    return { items, notes: t.notes };
  }

  function buildExcluir(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["ACTIVA", "FUENTE", "ID_SERVICIO_FUENTE"], problems);
    const items = t.records.map((r) => {
      const item = {
        activa: isActive(r),
        fuente: get(r, "FUENTE"),
        servicio: get(r, "SERVICIO_FUENTE"),
        servicioId: get(r, "ID_SERVICIO_FUENTE"),
        especialidad: get(r, "ESPECIALIDAD_FUENTE"),
        especialidadId: get(r, "ID_ESPECIALIDAD_FUENTE"),
        estudio: get(r, "NOMBRE_ESTUDIO"),
        estudioId: get(r, "ID_ESTUDIO_FUENTE"),
      };
      // Mismo criterio que explica la hoja EXCLUIR del Excel
      if (!isEmpty(item.estudioId)) item.alcance = "estudio";
      else if (!isEmpty(item.especialidadId)) item.alcance = "combinacion";
      else if (!isEmpty(item.servicioId)) item.alcance = "servicio";
      else item.alcance = "incompleta";
      return item;
    });
    return { items, notes: t.notes };
  }

  function buildBiogenomica(sheet, name, problems) {
    const t = readTable(sheet);
    checkColumns(t, name, ["NOMBRE DEL ESTUDIO", "ESPECIALIDAD", "SERVICIO"], problems);
    const items = t.records.map((r) => ({
      estudio: get(r, "NOMBRE DEL ESTUDIO"),
      especialidad: get(r, "ESPECIALIDAD"),
      servicio: get(r, "SERVICIO"),
      sgm: code(get(r, "COD SGM")),
      dm: code(get(r, "COD DM")),
      lm: code(get(r, "COD LM")),
      clipper: code(get(r, "COD CLIPPER")),
    }));
    return { items, notes: t.notes };
  }

  /*
   * Glosario: hoja con la lista de campos de la normalización. Se ubica la celda
   * "FUENTE" (encabezado de los procesos); a su izquierda está la descripción y,
   * dos columnas a la izquierda, la marca "*". Cada fila siguiente es un campo.
   */
  function buildGlosario(sheet, name, problems) {
    const rows = sheet ? sheet.rows : [];
    let headerRow = -1;
    let fieldCol = -1;
    for (let r = 0; r < Math.min(rows.length, 15) && headerRow < 0; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (norm(row[c]) === "FUENTE") {
          headerRow = r;
          fieldCol = c;
          break;
        }
      }
    }
    if (headerRow < 0) {
      if (sheet) problems.push(`En la hoja "${name}" no se encontró la celda "FUENTE" que encabeza el glosario.`);
      return { procesos: [], items: [], notes: [] };
    }
    const header = rows[headerRow];
    const procesos = [];
    for (let c = fieldCol + 1; c < header.length; c++) {
      if (isEmpty(header[c])) continue;
      const archivos = rows
        .slice(0, headerRow)
        .map((row) => clean((row || [])[c]))
        .filter((v) => typeof v === "string");
      procesos.push({ col: c, nombre: String(clean(header[c])), archivos });
    }
    const items = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const campo = clean(row[fieldCol]);
      if (isEmpty(campo)) continue;
      const descripcion = fieldCol >= 1 ? clean(row[fieldCol - 1]) : null;
      const marcado = fieldCol >= 2 ? norm(row[fieldCol - 2]) === "*" : false;
      const origen = {};
      for (const p of procesos) origen[p.nombre] = clean(row[p.col]);
      const textos = [descripcion, ...Object.values(origen)].filter((v) => typeof v === "string");
      items.push({
        campo: String(campo),
        descripcion,
        marcado,
        origen,
        pendiente: textos.some((s) => s.includes("?")),
        fila: r + 1,
      });
    }
    return { procesos, items, notes: [] };
  }

  // ---------- Entrada ----------

  function construirModelo(workbook, hojas) {
    const problems = [];
    const sh = (key) => requireSheet(workbook, hojas[key], problems);
    const modelo = {
      sedes: buildSedes(sh("sedes"), hojas.sedes, problems),
      especialidad: buildEspecialidadServicio(sh("especialidadServicio"), hojas.especialidadServicio, problems),
      unidad: buildUnidadNegocio(sh("unidadNegocio"), hojas.unidadNegocio, problems),
      estudios: buildEstudios(sh("estudios"), hojas.estudios, problems),
      biogenomica: buildBiogenomica(sh("biogenomica"), hojas.biogenomica, problems),
      excluir: buildExcluir(sh("excluir"), hojas.excluir, problems),
      glosario: buildGlosario(sh("glosario"), hojas.glosario, problems),
      problemas: problems,
    };
    return modelo;
  }

  global.Modelo = { construirModelo, norm };
})(window);
