/*
 * Tablero "Reglas de negocio consensuadas".
 * Arma la interfaz a partir del modelo (modelo.js) leído desde SharePoint (sharepoint.js)
 * o desde una copia local del Excel.
 */
(function () {
  "use strict";

  const CFG = window.APP_CONFIG || {};
  const app = document.getElementById("app");
  const sesionEl = document.getElementById("sesion");
  const archivoInput = document.getElementById("archivo-local");

  const state = {
    modelo: null,
    meta: null,
    tab: "sedes",
    filtros: {
      sedes: { estado: "falta", fuente: "", modelo: "", q: "" },
      especialidad: { estado: "todas", q: "" },
      unidad: { estado: "todas", q: "" },
      estudios: { estado: "todas", q: "" },
      biogenomica: { especialidad: "", codigos: "", q: "" },
      excluir: { estado: "todas", alcance: "", q: "" },
      glosario: { estado: "todos", q: "" },
    },
  };

  // ---------- Helpers de DOM (todo el texto entra como texto, nunca como HTML) ----------

  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === "class") el.className = v;
      else if (k.startsWith("on") && typeof v === "function") el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? "" : String(v));
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid === null || kid === undefined || kid === false) continue;
      el.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    }
    return el;
  }

  const norm = (s) =>
    String(s ?? "")
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .toLowerCase()
      .trim();

  function coincide(q, ...campos) {
    if (!q) return true;
    const needle = norm(q);
    return campos.some((c) => norm(fmt(c)).includes(needle));
  }

  function fmt(v) {
    if (v === null || v === undefined) return "";
    if (v instanceof Date) {
      const d = String(v.getUTCDate()).padStart(2, "0");
      const m = String(v.getUTCMonth() + 1).padStart(2, "0");
      return `${d}/${m}/${v.getUTCFullYear()}`;
    }
    return String(v);
  }

  function fechaHora(d) {
    if (!d) return "";
    return d.toLocaleString("es-AR", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23",
    });
  }

  function hora(d) {
    return d.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
  }

  const plural = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;

  // Solo se aceptan links https (evita enlaces javascript: u otros esquemas)
  function urlSegura(u) {
    try {
      const url = new URL(u);
      return url.protocol === "https:" ? url.href : null;
    } catch (_) {
      return null;
    }
  }

  function badge(texto, tipo) {
    return h("span", { class: `badge badge-${tipo}` }, texto);
  }

  function falta() {
    return h("span", { class: "falta" }, "Falta definir");
  }

  function mono(texto) {
    return h("span", { class: "mono" }, texto);
  }

  function ids(pares) {
    const partes = pares.filter(([, v]) => v !== null && v !== undefined && v !== "").map(([k, v]) => `${k} ${fmt(v)}`);
    return partes.length ? h("div", { class: "sub mono" }, partes.join(" · ")) : null;
  }

  function ruta(...partes) {
    const p = partes.filter((x) => x !== null && x !== undefined && x !== "");
    if (!p.length) return h("span", { class: "vacio" }, "—");
    return h("span", { class: "ruta" }, p.map((x, i) => [i ? h("span", { class: "sep", "aria-hidden": "true" }, " › ") : null, fmt(x)]));
  }

  function estadoRegla(activa) {
    return activa ? badge("Activa", "ok") : badge("Inactiva", "off");
  }

  function debounce(fn, ms) {
    let t;
    return (...a) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...a), ms);
    };
  }

  // ---------- Controles de filtro ----------

  function segmentado(opciones, actual, onChange, etiqueta) {
    const grupo = h("div", { class: "seg", role: "group", "aria-label": etiqueta });
    for (const op of opciones) {
      grupo.append(
        h(
          "button",
          {
            type: "button",
            class: op.value === actual ? "on" : "",
            "aria-pressed": op.value === actual ? "true" : "false",
            onclick: () => {
              for (const b of grupo.children) {
                b.classList.remove("on");
                b.setAttribute("aria-pressed", "false");
              }
              const btn = [...grupo.children][opciones.indexOf(op)];
              btn.classList.add("on");
              btn.setAttribute("aria-pressed", "true");
              onChange(op.value);
            },
          },
          op.label
        )
      );
    }
    return grupo;
  }

  function selector(etiqueta, opciones, actual, onChange) {
    const sel = h(
      "select",
      { onchange: (e) => onChange(e.target.value) },
      opciones.map((op) => h("option", { value: op.value, selected: op.value === actual }, op.label))
    );
    return h("label", { class: "campo" }, h("span", null, etiqueta), sel);
  }

  function buscador(placeholder, actual, onInput) {
    const input = h("input", {
      type: "search",
      placeholder,
      value: actual,
      "aria-label": placeholder,
      oninput: debounce((e) => onInput(e.target.value), 120),
    });
    return h("label", { class: "buscar" }, input);
  }

  function valoresUnicos(items, fn) {
    return [...new Set(items.map(fn).filter((v) => v !== null && v !== undefined && v !== ""))].sort((a, b) =>
      String(a).localeCompare(String(b), "es")
    );
  }

  function notasExcel(notas) {
    if (!notas || !notas.length) return null;
    return h(
      "details",
      { class: "notas" },
      h("summary", null, `Notas escritas en esta hoja del Excel (${notas.length})`),
      h("ul", null, notas.map((n) => h("li", null, n)))
    );
  }

  function tabla(columnas, filas, { vacio = "No hay filas con estos filtros.", clase = "" } = {}) {
    const thead = h("thead", null, h("tr", null, columnas.map((c) => h("th", { class: c.clase || null, scope: "col" }, c.titulo))));
    const tbody = h("tbody");
    if (!filas.length) {
      tbody.append(h("tr", { class: "fila-vacia" }, h("td", { colspan: columnas.length }, vacio)));
    }
    for (const fila of filas) tbody.append(fila);
    const alta = filas.length > 25 ? " alta" : "";
    return h("div", { class: `tabla-wrap ${clase}${alta}` }, h("table", null, thead, tbody));
  }

  function fila(celdas, clase) {
    return h("tr", { class: clase || null }, celdas);
  }

  function td(contenido, clase) {
    return h("td", { class: clase || null }, contenido);
  }

  function encabezadoSeccion(titulo, descripcion) {
    return h("div", { class: "seccion-head" }, h("h2", null, titulo), h("p", null, descripcion));
  }

  function contador(mostrando, total, unidad) {
    return h("p", { class: "contador" }, mostrando === total ? `${total} ${unidad}` : `Mostrando ${mostrando} de ${total} ${unidad}`);
  }

  // ---------- Pestañas ----------

  const TABS = [
    { id: "sedes", label: "Sedes", render: panelSedes },
    { id: "especialidad", label: "Especialidad → servicio", render: panelEspecialidad },
    { id: "unidad", label: "Cambios de Unidad de Negocio", render: panelUnidad },
    { id: "estudios", label: "Estudios que cambian de Servicio", render: panelEstudios },
    { id: "biogenomica", label: "Biogenómica", render: panelBiogenomica },
    { id: "excluir", label: "Excluir", render: panelExcluir },
    { id: "glosario", label: "Glosario", render: panelGlosario },
  ];

  function tabDesdeHash() {
    const id = location.hash.replace("#", "");
    return TABS.some((t) => t.id === id) ? id : null;
  }

  function irA(id, { scroll = false } = {}) {
    state.tab = id;
    history.replaceState(null, "", "#" + id);
    document.querySelectorAll(".tabs [role=tab]").forEach((b) => {
      const on = b.dataset.tab === id;
      b.setAttribute("aria-selected", on ? "true" : "false");
      b.tabIndex = on ? 0 : -1;
      if (on) b.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
    document.querySelectorAll(".kpi").forEach((k) => k.classList.toggle("on", k.dataset.tab === id));
    renderPanel();
    if (scroll) document.querySelector(".tabs").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // ---------- Panel: Sedes ----------

  function panelSedes() {
    const f = state.filtros.sedes;
    const todas = state.modelo.sedes.items;
    const resultados = h("div");

    const porFuente = new Map();
    for (const s of todas) {
      if (!porFuente.has(s.fuente)) porFuente.set(s.fuente, { total: 0, completas: 0 });
      const g = porFuente.get(s.fuente);
      g.total++;
      if (s.completa) g.completas++;
    }

    function actualizar() {
      const visibles = todas.filter(
        (s) =>
          (f.estado === "todas" || (f.estado === "falta" ? !s.completa : s.completa)) &&
          (!f.fuente || s.fuente === f.fuente) &&
          (!f.modelo || norm(s.campos.MODELO_NEGOCIO) === norm(f.modelo)) &&
          coincide(f.q, s.nombre, s.campos.NOMBRE_SEDE_PARSEADO, s.codigo, s.origenId, s.numero, s.observaciones)
      );
      const fuentes = [...new Set(visibles.map((s) => s.fuente))].sort((a, b) => a.localeCompare(b, "es"));
      const filas = [];
      for (const fuente of fuentes) {
        const g = porFuente.get(fuente);
        filas.push(
          h(
            "tr",
            { class: "grupo" },
            h("td", { colspan: 8 }, h("span", { class: "grupo-nombre" }, fuente), h("span", { class: "grupo-info" }, `${g.completas} de ${g.total} completas`))
          )
        );
        for (const s of visibles.filter((x) => x.fuente === fuente)) {
          const dato = (key) => (s.faltan.length && !esDefinido(s.campos[key]) ? td(falta(), "td-falta") : td(fmt(s.campos[key])));
          filas.push(
            fila([
              td(s.completa ? badge("Completa", "ok") : badge(`Faltan ${s.faltan.length}`, "warn")),
              td([h("strong", null, fmt(s.nombre)), ids([["origen", s.origenId ?? "—"], ["código", s.codigo ?? "—"]])]),
              dato("NOMBRE_SEDE_PARSEADO"),
              dato("TIPO_SEDE"),
              dato("EMPRESA"),
              dato("MODALIDAD"),
              dato("MODELO_NEGOCIO"),
              td(fmt(s.observaciones), "obs"),
            ])
          );
        }
      }
      resultados.replaceChildren(
        contador(visibles.length, todas.length, "sedes"),
        tabla(
          [
            { titulo: "Estado" },
            { titulo: "Sede en la fuente" },
            { titulo: "Sede normalizada" },
            { titulo: "Tipo de sede" },
            { titulo: "Empresa" },
            { titulo: "Modalidad" },
            { titulo: "Modelo de negocio" },
            { titulo: "Observaciones" },
          ],
          filas,
          { clase: "tabla-sedes" }
        )
      );
    }

    const fuentes = valoresUnicos(todas, (s) => s.fuente);
    const modelos = valoresUnicos(todas, (s) => (s.campos.MODELO_NEGOCIO ? String(s.campos.MODELO_NEGOCIO).trim() : null));
    const nodo = h(
      "div",
      null,
      encabezadoSeccion(
        "Sedes",
        "Cada origen de cada sistema se traduce a una sede normalizada con su tipo, empresa, modalidad y modelo de negocio. Una sede está completa cuando tiene definidos los cinco datos: sede normalizada, tipo de sede, empresa, modalidad y modelo de negocio."
      ),
      h(
        "div",
        { class: "filtros" },
        segmentado(
          [
            { value: "todas", label: "Todas" },
            { value: "falta", label: "Falta definir" },
            { value: "completa", label: "Completas" },
          ],
          f.estado,
          (v) => ((f.estado = v), actualizar()),
          "Estado de la sede"
        ),
        selector("Fuente", [{ value: "", label: "Todas" }, ...fuentes.map((x) => ({ value: x, label: x }))], f.fuente, (v) => ((f.fuente = v), actualizar())),
        selector("Modelo", [{ value: "", label: "Todos" }, ...modelos.map((x) => ({ value: x, label: x }))], f.modelo, (v) => ((f.modelo = v), actualizar())),
        buscador("Buscar sede, código u origen", f.q, (v) => ((f.q = v), actualizar()))
      ),
      h("p", { class: "leyenda" }, falta(), " dato que todavía no se consensuó con los pilares"),
      resultados,
      notasExcel(state.modelo.sedes.notes)
    );
    actualizar();
    return nodo;
  }

  function esDefinido(v) {
    if (v === null || v === undefined || String(v).trim() === "") return false;
    const s = norm(v);
    return s !== "-" && !s.includes("definir");
  }

  // ---------- Paneles de reglas ----------

  function filtroEstado(f, actualizar) {
    return segmentado(
      [
        { value: "todas", label: "Todas" },
        { value: "activas", label: "Activas" },
        { value: "inactivas", label: "Inactivas" },
      ],
      f.estado,
      (v) => ((f.estado = v), actualizar()),
      "Estado de la regla"
    );
  }

  function pasaEstado(f, item) {
    return f.estado === "todas" || (f.estado === "activas" ? item.activa : !item.activa);
  }

  function panelEspecialidad() {
    const f = state.filtros.especialidad;
    const items = state.modelo.especialidad.items;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (r) => pasaEstado(f, r) && coincide(f.q, r.fuente, r.servicio, r.especialidad, r.servicioNuevo, r.especialidadNueva, r.servicioId, r.especialidadId, r.servicioNuevoId, r.especialidadNuevaId)
      );
      resultados.replaceChildren(
        contador(visibles.length, items.length, "reglas"),
        tabla(
          [{ titulo: "Estado" }, { titulo: "Fuente" }, { titulo: "Hoy en la fuente" }, { titulo: "", clase: "flecha" }, { titulo: "Pasa a" }],
          visibles.map((r) =>
            fila(
              [
                td(estadoRegla(r.activa)),
                td(mono(fmt(r.fuente))),
                td([ruta(r.servicio, r.especialidad), ids([["servicio", r.servicioId], ["especialidad", r.especialidadId]])]),
                td("→", "flecha"),
                td([ruta(r.servicioNuevo, r.especialidadNueva), ids([["servicio", r.servicioNuevoId], ["especialidad", r.especialidadNuevaId]])], "destino"),
              ],
              r.activa ? "" : "inactiva"
            )
          )
        )
      );
    }
    const nodo = h(
      "div",
      null,
      encabezadoSeccion(
        "Especialidad → servicio",
        "Especialidades que, al normalizar, pasan a otro servicio o quedan como un servicio propio. La regla se aplica a la combinación de servicio y especialidad de la fuente."
      ),
      h("div", { class: "filtros" }, filtroEstado(f, actualizar), buscador("Buscar servicio, especialidad o ID", f.q, (v) => ((f.q = v), actualizar()))),
      resultados,
      notasExcel(state.modelo.especialidad.notes)
    );
    actualizar();
    return nodo;
  }

  function panelUnidad() {
    const f = state.filtros.unidad;
    const items = state.modelo.unidad.items;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (r) => pasaEstado(f, r) && coincide(f.q, r.fuente, r.servicio, r.especialidad, r.unidadActual, r.unidadNueva, r.servicioId, r.especialidadId)
      );
      resultados.replaceChildren(
        contador(visibles.length, items.length, "reglas"),
        tabla(
          [{ titulo: "Estado" }, { titulo: "Fuente" }, { titulo: "Servicio › especialidad" }, { titulo: "Unidad de negocio" }],
          visibles.map((r) =>
            fila(
              [
                td(estadoRegla(r.activa)),
                td(mono(fmt(r.fuente))),
                td([ruta(r.servicio, r.especialidad), ids([["servicio", r.servicioId], ["especialidad", r.especialidadId]])]),
                td(h("span", { class: "cambio" }, h("span", { class: "chip" }, fmt(r.unidadActual) || "—"), h("span", { class: "flecha" }, "→"), h("span", { class: "chip chip-accent" }, fmt(r.unidadNueva) || "—"))),
              ],
              r.activa ? "" : "inactiva"
            )
          )
        )
      );
    }
    const nodo = h(
      "div",
      null,
      encabezadoSeccion("Cambios de Unidad de Negocio", "Combinaciones de servicio y especialidad que, al normalizar, pasan de una unidad de negocio a otra."),
      h("div", { class: "filtros" }, filtroEstado(f, actualizar), buscador("Buscar servicio, especialidad o unidad", f.q, (v) => ((f.q = v), actualizar()))),
      resultados,
      notasExcel(state.modelo.unidad.notes)
    );
    actualizar();
    return nodo;
  }

  function panelEstudios() {
    const f = state.filtros.estudios;
    const items = state.modelo.estudios.items;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (r) => pasaEstado(f, r) && coincide(f.q, r.fuente, r.estudio, r.praId, r.praCodigo, r.especialidad, r.especialidadNueva, r.servicioNuevo)
      );
      resultados.replaceChildren(
        contador(visibles.length, items.length, "reglas"),
        tabla(
          [{ titulo: "Estado" }, { titulo: "Fuente" }, { titulo: "Estudio" }, { titulo: "Especialidad en la fuente" }, { titulo: "", clase: "flecha" }, { titulo: "Pasa a" }],
          visibles.map((r) =>
            fila(
              [
                td(estadoRegla(r.activa)),
                td(mono(fmt(r.fuente))),
                td([h("strong", null, fmt(r.estudio)), ids([["id", r.praId], ["código", r.praCodigo], ["región", r.praRegion]])]),
                td([ruta(r.especialidad), ids([["id", r.especialidadId]])]),
                td("→", "flecha"),
                td([ruta(r.servicioNuevo, r.especialidadNueva), ids([["servicio", r.servicioNuevoId], ["especialidad", r.especialidadNuevaId]])], "destino"),
              ],
              r.activa ? "" : "inactiva"
            )
          )
        )
      );
    }
    const nodo = h(
      "div",
      null,
      encabezadoSeccion(
        "Estudios que cambian de Servicio",
        "Estudios puntuales que se reasignan a otro servicio y especialidad, sin importar en qué especialidad estén cargados en la fuente."
      ),
      h("div", { class: "filtros" }, filtroEstado(f, actualizar), buscador("Buscar estudio, código o especialidad", f.q, (v) => ((f.q = v), actualizar()))),
      resultados,
      notasExcel(state.modelo.estudios.notes)
    );
    actualizar();
    return nodo;
  }

  const ALCANCE = {
    servicio: { label: "Todo el servicio", tipo: "accent" },
    combinacion: { label: "Servicio + especialidad", tipo: "info" },
    estudio: { label: "Un estudio", tipo: "info" },
    incompleta: { label: "Sin ID", tipo: "warn" },
  };

  function panelExcluir() {
    const f = state.filtros.excluir;
    const items = state.modelo.excluir.items;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (r) => pasaEstado(f, r) && (!f.alcance || r.alcance === f.alcance) && coincide(f.q, r.fuente, r.servicio, r.especialidad, r.estudio, r.servicioId, r.especialidadId, r.estudioId)
      );
      resultados.replaceChildren(
        contador(visibles.length, items.length, "exclusiones"),
        tabla(
          [{ titulo: "Estado" }, { titulo: "Fuente" }, { titulo: "Alcance" }, { titulo: "Qué se excluye" }],
          visibles.map((r) => {
            const a = ALCANCE[r.alcance];
            const que =
              r.alcance === "estudio"
                ? [h("strong", null, fmt(r.estudio) || `Estudio ${fmt(r.estudioId)}`), ids([["estudio", r.estudioId], ["servicio", r.servicioId], ["especialidad", r.especialidadId]])]
                : r.alcance === "combinacion"
                ? [ruta(r.servicio, r.especialidad), ids([["servicio", r.servicioId], ["especialidad", r.especialidadId]])]
                : [ruta(r.servicio), ids([["servicio", r.servicioId]])];
            return fila([td(estadoRegla(r.activa)), td(mono(fmt(r.fuente))), td(badge(a.label, a.tipo)), td(que)], r.activa ? "" : "inactiva");
          })
        )
      );
    }
    const nodo = h(
      "div",
      null,
      encabezadoSeccion(
        "Excluir",
        "Servicios, combinaciones de servicio y especialidad, o estudios que se excluyen de la carga en los QVD de normalización, por desuso, por ser de prueba u otros motivos."
      ),
      h(
        "div",
        { class: "filtros" },
        filtroEstado(f, actualizar),
        selector(
          "Alcance",
          [{ value: "", label: "Todos" }, ...["servicio", "combinacion", "estudio"].map((k) => ({ value: k, label: ALCANCE[k].label }))],
          f.alcance,
          (v) => ((f.alcance = v), actualizar())
        ),
        buscador("Buscar servicio, especialidad o estudio", f.q, (v) => ((f.q = v), actualizar()))
      ),
      resultados,
      notasExcel(state.modelo.excluir.notes)
    );
    actualizar();
    return nodo;
  }

  // ---------- Panel: Biogenómica ----------

  const SISTEMAS = [
    { key: "sgm", label: "SGM" },
    { key: "dm", label: "DM" },
    { key: "lm", label: "LM" },
    { key: "clipper", label: "Clipper" },
  ];

  function panelBiogenomica() {
    const f = state.filtros.biogenomica;
    const items = state.modelo.biogenomica.items;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (r) =>
          (!f.especialidad || r.especialidad === f.especialidad) &&
          (!f.codigos || !r[f.codigos]) &&
          coincide(f.q, r.estudio, r.sgm, r.dm, r.lm, r.clipper, r.especialidad)
      );
      const codigo = (v) => (v ? td(mono(v), "codigo") : td(h("span", { class: "vacio" }, "—"), "codigo"));
      resultados.replaceChildren(
        contador(visibles.length, items.length, "prestaciones"),
        tabla(
          [{ titulo: "Estudio" }, { titulo: "Especialidad" }, { titulo: "Servicio" }, ...SISTEMAS.map((s) => ({ titulo: s.label, clase: "codigo" }))],
          visibles.map((r) => fila([td(h("strong", null, fmt(r.estudio))), td(fmt(r.especialidad)), td(fmt(r.servicio)), codigo(r.sgm), codigo(r.dm), codigo(r.lm), codigo(r.clipper)]))
        )
      );
    }
    const cobertura = h(
      "div",
      { class: "cobertura" },
      SISTEMAS.map((s) => {
        const n = items.filter((r) => r[s.key]).length;
        return h("div", { class: "cob" }, h("span", { class: "cob-label" }, `Con código ${s.label}`), h("span", { class: "cob-valor" }, `${n}`, h("small", null, ` de ${items.length}`)));
      })
    );
    const especialidades = valoresUnicos(items, (r) => r.especialidad);
    const nodo = h(
      "div",
      null,
      encabezadoSeccion("Biogenómica", "Prestaciones de genética mapeadas entre sistemas: cada estudio con su código en SGM, DM, LM y Clipper."),
      cobertura,
      h(
        "div",
        { class: "filtros" },
        selector("Especialidad", [{ value: "", label: "Todas" }, ...especialidades.map((x) => ({ value: x, label: x }))], f.especialidad, (v) => ((f.especialidad = v), actualizar())),
        selector(
          "Códigos",
          [{ value: "", label: "Todos" }, ...SISTEMAS.filter((s) => s.key !== "clipper").map((s) => ({ value: s.key, label: `Sin código ${s.label}` }))],
          f.codigos,
          (v) => ((f.codigos = v), actualizar())
        ),
        buscador("Buscar estudio o código", f.q, (v) => ((f.q = v), actualizar()))
      ),
      resultados,
      notasExcel(state.modelo.biogenomica.notes)
    );
    actualizar();
    return nodo;
  }

  // ---------- Panel: Glosario ----------

  function panelGlosario() {
    const f = state.filtros.glosario;
    const { items, procesos } = state.modelo.glosario;
    const resultados = h("div");
    function actualizar() {
      const visibles = items.filter(
        (g) =>
          (f.estado === "todos" || (f.estado === "pendientes" ? g.pendiente : g.marcado)) &&
          coincide(f.q, g.campo, g.descripcion, ...Object.values(g.origen))
      );
      resultados.replaceChildren(
        contador(visibles.length, items.length, "términos"),
        tabla(
          [
            { titulo: "Campo" },
            { titulo: "Descripción / ejemplo" },
            ...procesos.map((p) => ({ titulo: p.nombre, clase: "proceso" })),
          ],
          visibles.map((g) =>
            fila(
              [
                td([
                  h("span", { class: "mono campo-nombre" }, g.campo),
                  g.marcado ? h("span", { class: "marca", title: "Marcado con * en el Excel" }, "*") : null,
                  g.pendiente ? h("div", null, badge("Pendiente de revisión", "warn")) : null,
                ]),
                td(g.descripcion ? fmt(g.descripcion) : h("span", { class: "vacio" }, "—"), "desc"),
                ...procesos.map((p) => td(g.origen[p.nombre] ? fmt(g.origen[p.nombre]) : h("span", { class: "vacio" }, "—"), "proceso")),
              ],
              g.pendiente ? "pendiente" : ""
            )
          ),
          { clase: "tabla-glosario" }
        )
      );
    }
    const archivos = procesos.filter((p) => p.archivos.length);
    const nodo = h(
      "div",
      null,
      encabezadoSeccion(
        "Glosario",
        "Campos del modelo normalizado y de dónde sale cada uno en cada proceso de normalización de Qlik. Un signo de pregunta en el Excel marca el término como pendiente de revisión."
      ),
      archivos.length
        ? h(
            "details",
            { class: "notas" },
            h("summary", null, "Archivos de Qlik por proceso"),
            h("ul", null, archivos.map((p) => h("li", null, h("strong", null, p.nombre + ": "), mono(p.archivos.join(" · ")))))
          )
        : null,
      h(
        "div",
        { class: "filtros" },
        segmentado(
          [
            { value: "todos", label: "Todos" },
            { value: "pendientes", label: "Pendientes" },
            { value: "marcados", label: "Marcados (*)" },
          ],
          f.estado,
          (v) => ((f.estado = v), actualizar()),
          "Filtro de términos"
        ),
        buscador("Buscar campo o descripción", f.q, (v) => ((f.q = v), actualizar()))
      ),
      resultados
    );
    actualizar();
    return nodo;
  }

  // ---------- Tablero ----------

  function contarReglas(items) {
    const activas = items.filter((r) => r.activa).length;
    return { activas, inactivas: items.length - activas };
  }

  function kpis() {
    const m = state.modelo;
    const sedes = m.sedes.items;
    const completas = sedes.filter((s) => s.completa).length;
    const esp = contarReglas(m.especialidad.items);
    const uni = contarReglas(m.unidad.items);
    const est = contarReglas(m.estudios.items);
    const exc = contarReglas(m.excluir.items);
    const pendientes = m.glosario.items.filter((g) => g.pendiente).length;
    const tarjetas = [
      {
        tab: "sedes",
        label: "Sedes",
        valor: `${completas}/${sedes.length}`,
        sub: `${sedes.length - completas} con datos por definir`,
        progreso: sedes.length ? completas / sedes.length : 0,
      },
      { tab: "especialidad", label: "Especialidad → servicio", valor: esp.activas, sub: `${esp.activas === 1 ? "activa" : "activas"} · ${plural(esp.inactivas, "inactiva", "inactivas")}` },
      { tab: "unidad", label: "Cambios de Unidad de Negocio", valor: uni.activas, sub: `${uni.activas === 1 ? "activa" : "activas"} · ${plural(uni.inactivas, "inactiva", "inactivas")}` },
      { tab: "estudios", label: "Estudios que cambian de Servicio", valor: est.activas, sub: `${est.activas === 1 ? "activa" : "activas"} · ${plural(est.inactivas, "inactiva", "inactivas")}` },
      { tab: "biogenomica", label: "Biogenómica", valor: m.biogenomica.items.length, sub: "prestaciones mapeadas (DM · LM · Clipper)" },
      { tab: "excluir", label: "Excluir", valor: exc.activas, sub: `${exc.activas === 1 ? "exclusión activa" : "exclusiones activas"} · ${plural(exc.inactivas, "inactiva", "inactivas")}` },
      { tab: "glosario", label: "Glosario", valor: m.glosario.items.length, sub: `términos · ${plural(pendientes, "pendiente", "pendientes")} de revisión` },
    ];
    return h(
      "div",
      { class: "kpis" },
      tarjetas.map((t) => {
        const barra = t.progreso !== undefined ? h("div", { class: "barra", role: "img", "aria-label": `${Math.round(t.progreso * 100)}% completas` }, h("span")) : null;
        if (barra) barra.firstChild.style.width = `${(t.progreso * 100).toFixed(1)}%`;
        return h(
          "button",
          { type: "button", class: "kpi" + (state.tab === t.tab ? " on" : ""), "data-tab": t.tab, onclick: () => irA(t.tab, { scroll: true }) },
          h("span", { class: "kpi-label" }, t.label),
          h("span", { class: "kpi-valor" }, String(t.valor)),
          barra,
          h("span", { class: "kpi-sub" }, t.sub)
        );
      })
    );
  }

  function pestañas() {
    const nav = h("div", { class: "tabs", role: "tablist", "aria-label": "Secciones" });
    TABS.forEach((t) => {
      nav.append(
        h(
          "button",
          {
            type: "button",
            role: "tab",
            "data-tab": t.id,
            "aria-selected": state.tab === t.id ? "true" : "false",
            "aria-controls": "panel",
            tabindex: state.tab === t.id ? "0" : "-1",
            onclick: () => irA(t.id),
          },
          t.label
        )
      );
    });
    nav.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
      const i = TABS.findIndex((t) => t.id === state.tab);
      const j = (i + (e.key === "ArrowRight" ? 1 : TABS.length - 1)) % TABS.length;
      irA(TABS[j].id);
      nav.querySelector(`[data-tab="${TABS[j].id}"]`).focus();
    });
    return nav;
  }

  function lineaFuente() {
    const meta = state.meta;
    if (meta.origen === "local") {
      return h(
        "p",
        { class: "fuente" },
        "Fuente: ",
        mono(meta.nombre),
        ` (copia local, modificada el ${fechaHora(meta.modificado)}). `,
        h("button", { type: "button", class: "link", onclick: () => archivoInput.click() }, "Abrir otro archivo")
      );
    }
    return h(
      "p",
      { class: "fuente" },
      "Fuente: ",
      urlSegura(meta.webUrl) ? h("a", { href: urlSegura(meta.webUrl), target: "_blank", rel: "noopener" }, mono(meta.nombre)) : mono(meta.nombre),
      " en SharePoint",
      meta.modificado ? `, modificado el ${fechaHora(meta.modificado)}` : "",
      meta.modificadoPor ? ` por ${meta.modificadoPor}` : "",
      `. Leído a las ${hora(meta.leido)}. `,
      h("button", { type: "button", class: "link", onclick: () => cargarDesdeSharePoint({ silencioso: true }) }, "Actualizar")
    );
  }

  function cajaProponer() {
    return h(
      "aside",
      { class: "proponer" },
      h("p", { class: "proponer-titulo" }, "Para proponer cambios"),
      urlSegura(CFG.maestroUrl)
        ? h(
            "a",
            { class: "btn-maestro", href: urlSegura(CFG.maestroUrl), target: "_blank", rel: "noopener" },
            h("span", { "aria-hidden": "true" }, "✎"),
            " Abrir MAESTRO - GOBIERNO DE DATOS ",
            h("span", { "aria-hidden": "true" }, "↗")
          )
        : null,
      h(
        "ol",
        null,
        h("li", null, h("strong", null, "Pilares"), " editan el MAESTRO en SharePoint"),
        h("li", null, h("strong", null, "BI"), " vuelca los acuerdos en Reglas de Negocio"),
        h("li", null, h("strong", null, "Qlik"), " consume Reglas de Negocio")
      )
    );
  }

  function renderTablero() {
    const m = state.modelo;
    const avisos =
      m.problemas.length > 0
        ? h(
            "div",
            { class: "aviso aviso-warn", role: "alert" },
            h("strong", null, "Revisar el Excel: "),
            "algunas partes no se pudieron leer y pueden aparecer vacías.",
            h("ul", null, m.problemas.map((p) => h("li", null, p)))
          )
        : null;
    const avisoLocal =
      state.meta.origen === "local"
        ? h("div", { class: "aviso aviso-info" }, "Estás viendo una copia local del Excel, no la versión en SharePoint.")
        : null;

    const bloques = [
      h(
        "div",
        { class: "hero" },
        h(
          "div",
          { class: "hero-texto" },
          h("p", { class: "eyebrow" }, "Gobierno de Datos · Diagnóstico Maipú"),
          h("h1", null, "Reglas de negocio consensuadas"),
          h("p", { class: "bajada" }, "Lo que ya está acordado con los pilares, escrito en palabras, y lo que falta definir."),
          lineaFuente()
        ),
        cajaProponer()
      ),
      avisoLocal,
      avisos,
      kpis(),
      pestañas(),
      h("section", { id: "panel", class: "panel", role: "tabpanel" }),
      h(
        "footer",
        { class: "pie" },
        state.meta.origen === "local"
          ? "Datos de una copia local del Excel."
          : "Los datos se leen en vivo desde SharePoint con tu usuario. Ves solo lo que tu cuenta tiene permiso de leer."
      ),
    ];
    app.replaceChildren(...bloques.filter(Boolean));
    renderPanel();
  }

  function renderPanel() {
    const panel = document.getElementById("panel");
    if (!panel || !state.modelo) return;
    const tab = TABS.find((t) => t.id === state.tab) || TABS[0];
    panel.setAttribute("aria-label", tab.label);
    panel.replaceChildren(tab.render());
  }

  // ---------- Pantallas de estado ----------

  function pantalla({ eyebrow, titulo, texto, acciones = [], detalle, tipo = "" }) {
    app.replaceChildren(
      h(
        "div",
        { class: `pantalla ${tipo}` },
        h("p", { class: "eyebrow" }, eyebrow || "Gobierno de Datos · Diagnóstico Maipú"),
        h("h1", null, titulo),
        texto ? h("p", { class: "bajada" }, texto) : null,
        acciones.length ? h("div", { class: "acciones" }, acciones) : null,
        detalle ? h("details", { class: "detalle" }, h("summary", null, "Detalle técnico"), h("pre", null, detalle)) : null
      )
    );
  }

  function cargando(texto) {
    app.replaceChildren(h("div", { class: "pantalla" }, h("div", { class: "spinner", "aria-hidden": "true" }), h("p", { class: "estado-carga" }, texto)));
  }

  function botonLocal(texto = "Abrir una copia local del Excel") {
    return h("button", { type: "button", class: "btn btn-sec", onclick: () => archivoInput.click() }, texto);
  }

  // La copia local solo se ofrece al probar en la propia computadora (localhost).
  // En el sitio publicado, la única forma de ver el tablero es iniciar sesión.
  function esLocalhost() {
    return ["localhost", "127.0.0.1", "[::1]"].indexOf(location.hostname) >= 0;
  }

  function renderIngreso() {
    pantalla({
      titulo: "Reglas de negocio consensuadas",
      texto: "Ingresá con tu cuenta de Microsoft de la empresa para ver las reglas. Los datos se leen en vivo desde SharePoint.",
      acciones: [
        h("button", { type: "button", class: "btn", onclick: () => SharePoint.ingresar().catch(renderError) }, "Ingresar con Microsoft"),
        new URLSearchParams(location.search).has("archivo") && esLocalhost() ? botonLocal() : null,
      ],
    });
  }

  function renderSinConfigurar() {
    pantalla({
      titulo: "Falta conectar con Microsoft 365",
      texto:
        esLocalhost()
          ? "Todavía no se cargaron tenantId y clientId en config.js (ver README, pasos 3 y 4). Mientras tanto, podés probar el tablero con una copia local del Excel: se lee en tu navegador y no se sube a ningún lado."
          : "El tablero todavía no está conectado con las cuentas de Microsoft de la empresa. Va a estar disponible en cuanto se complete la configuración.",
      acciones: esLocalhost() ? [botonLocal()] : [],
    });
  }

  function renderError(err) {
    console.error("[Reglas de negocio]", err);
    const tipo = err && err.tipo;
    const textos = {
      "sin-acceso": [
        "No tenés acceso al Excel en SharePoint",
        "Pedí permiso de lectura sobre DataGob_Reglas_Negocio.xlsx (sitio Qlik-Reportes) a quien administra el sitio y volvé a intentar.",
      ],
      "no-encontrado": ["No se encontró el Excel en SharePoint", err.detalle || "Revisá la configuración de la ubicación en config.js."],
      auth: ["No se pudo iniciar sesión", "Probá ingresar de nuevo. Si el problema sigue, avisale a IT con el detalle técnico."],
      formato: ["No se pudo leer el Excel", "El archivo no tiene el formato esperado. Avisale a BI con el detalle técnico."],
      config: ["Falta configuración", err.message],
    };
    const [titulo, texto] = textos[tipo] || ["No se pudo cargar el tablero", "Probá de nuevo en unos minutos. Si el problema sigue, avisale a IT con el detalle técnico."];
    const esMsal = err && err.errorCode;
    pantalla({
      titulo,
      texto,
      tipo: "pantalla-error",
      acciones: [
        h("button", { type: "button", class: "btn", onclick: () => location.reload() }, "Reintentar"),
        esMsal ? h("button", { type: "button", class: "btn btn-sec", onclick: () => SharePoint.ingresar() }, "Ingresar de nuevo") : null,
      ],
      detalle: [err && err.message, err && err.detalle, esMsal ? `errorCode: ${err.errorCode}` : null, err && err.errorMessage].filter(Boolean).join("\n"),
    });
  }

  function renderSesion() {
    sesionEl.replaceChildren();
    if (state.meta && state.meta.origen === "local") {
      sesionEl.append(h("span", { class: "chip" }, "Copia local"));
      return;
    }
    const u = window.SharePoint && SharePoint.usuario && SharePoint.estaConfigurado() ? SharePoint.usuario() : null;
    if (!u) return;
    const iniciales = u.nombre
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((p) => p[0].toUpperCase())
      .join("");
    sesionEl.append(
      h("span", { class: "avatar", "aria-hidden": "true" }, iniciales),
      h("span", { class: "usuario", title: u.email }, u.nombre),
      h("button", { type: "button", class: "link", onclick: () => SharePoint.salir() }, "Salir")
    );
  }

  // ---------- Carga de datos ----------

  async function procesar(buffer, meta) {
    let wb;
    try {
      wb = await XlsxReader.readWorkbook(buffer);
    } catch (err) {
      const e = new Error(err.message);
      e.tipo = "formato";
      throw e;
    }
    state.modelo = Modelo.construirModelo(wb, CFG.hojas || {});
    state.meta = meta;
    renderSesion();
    renderTablero();
  }

  async function cargarDesdeSharePoint({ silencioso = false } = {}) {
    const scrollY = window.scrollY;
    if (!silencioso || !state.modelo) cargando("Leyendo el Excel en SharePoint…");
    try {
      const { buffer, meta } = await SharePoint.leerExcel();
      await procesar(buffer, { ...meta, origen: "sharepoint", leido: new Date() });
      if (silencioso) window.scrollTo(0, scrollY);
    } catch (err) {
      renderError(err);
    }
  }

  archivoInput.addEventListener("change", async () => {
    const file = archivoInput.files && archivoInput.files[0];
    archivoInput.value = "";
    if (!file) return;
    cargando("Leyendo " + file.name + "…");
    try {
      const buffer = await file.arrayBuffer();
      await procesar(buffer, { origen: "local", nombre: file.name, modificado: new Date(file.lastModified), leido: new Date() });
    } catch (err) {
      renderError(err);
    }
  });

  window.addEventListener("hashchange", () => {
    const id = tabDesdeHash();
    if (id && id !== state.tab && state.modelo) irA(id);
  });

  async function iniciar() {
    state.tab = tabDesdeHash() || "sedes";
    if (!window.SharePoint || !SharePoint.estaConfigurado()) {
      renderSinConfigurar();
      return;
    }
    try {
      cargando("Conectando con Microsoft 365…");
      const cuenta = await SharePoint.iniciar();
      renderSesion();
      if (!cuenta) {
        renderIngreso();
        return;
      }
      await cargarDesdeSharePoint();
    } catch (err) {
      if (err && err.errorCode && !err.tipo) err.tipo = "auth";
      renderError(err);
    }
  }

  iniciar();
})();
