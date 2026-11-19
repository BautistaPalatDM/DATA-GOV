/*
 * Conexión con Microsoft 365: inicio de sesión (MSAL) y lectura del Excel
 * desde SharePoint con Microsoft Graph.
 *
 * Los permisos son DELEGADOS: la app lee el archivo como el usuario que inició
 * sesión. Si esa persona no tiene acceso al Excel en SharePoint, no ve los datos.
 */
(function (global) {
  "use strict";

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Files.Read.All", "Sites.Read.All"];
  const LOCATION_KEY = "reglas-negocio.ubicacion";

  let pca = null;
  let account = null;

  function config() {
    return global.APP_CONFIG || {};
  }

  function estaConfigurado() {
    const c = config();
    const ok = (v) => typeof v === "string" && /^[0-9a-f-]{36}$/i.test(v.trim());
    return ok(c.tenantId) && ok(c.clientId);
  }

  class ErrorConexion extends Error {
    constructor(message, { tipo = "general", detalle = "" } = {}) {
      super(message);
      this.tipo = tipo;
      this.detalle = detalle;
    }
  }

  // ---------- Sesión ----------

  async function iniciar() {
    const c = config();
    const base = new URL("./", global.location.href).href;
    pca = new msal.PublicClientApplication({
      auth: {
        clientId: c.clientId.trim(),
        authority: "https://login.microsoftonline.com/" + c.tenantId.trim(),
        redirectUri: new URL("redirect.html", base).href,
        postLogoutRedirectUri: base,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
    await pca.initialize();
    const result = await pca.handleRedirectPromise();
    account = (result && result.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
    return account;
  }

  function usuario() {
    return account ? { nombre: account.name || account.username, email: account.username } : null;
  }

  function ingresar() {
    return pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  }

  function salir() {
    try {
      sessionStorage.removeItem(LOCATION_KEY);
    } catch (_) {}
    return pca.logoutRedirect({ account });
  }

  async function token() {
    try {
      const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
      return r.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        await pca.acquireTokenRedirect({ scopes: SCOPES, account });
        return new Promise(() => {}); // la página navega a Microsoft
      }
      throw new ErrorConexion("No se pudo obtener permiso para leer SharePoint.", {
        tipo: "auth",
        detalle: err.errorCode || err.message,
      });
    }
  }

  // ---------- Graph ----------

  async function graph(path, accessToken) {
    const res = await fetch(GRAPH + path, { headers: { Authorization: "Bearer " + accessToken } });
    if (res.ok) return res.json();
    let detalle = res.status + " " + res.statusText;
    try {
      const body = await res.json();
      if (body && body.error) detalle = res.status + " · " + body.error.code + " · " + body.error.message;
    } catch (_) {}
    const err = new ErrorConexion("Error al consultar SharePoint.", { tipo: "graph", detalle });
    err.status = res.status;
    throw err;
  }

  // Codificación de una URL de SharePoint para el endpoint /shares de Graph
  function shareId(url) {
    const b64 = btoa(String.fromCharCode(...new TextEncoder().encode(url)));
    return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }

  function guardarUbicacion(loc) {
    try {
      sessionStorage.setItem(LOCATION_KEY, JSON.stringify(loc));
    } catch (_) {}
    console.info(
      "[Reglas de negocio] Ubicación del Excel en SharePoint. Para fijarla, copiá estos valores en config.js:",
      "\n  driveId:", loc.driveId,
      "\n  itemId: ", loc.itemId
    );
    return loc;
  }

  /*
   * Encuentra el archivo en SharePoint. Orden:
   *   1. driveId + itemId fijados en config.js (lo más rápido y estable)
   *   2. ubicación encontrada antes en esta sesión
   *   3. link del archivo (fileUrl) con el endpoint /shares
   *   4. búsqueda por nombre en las bibliotecas del sitio
   */
  async function ubicarArchivo(accessToken) {
    const sp = config().sharepoint || {};
    if (sp.driveId && sp.itemId) return { driveId: sp.driveId.trim(), itemId: sp.itemId.trim() };

    try {
      const cached = JSON.parse(sessionStorage.getItem(LOCATION_KEY) || "null");
      if (cached && cached.driveId && cached.itemId) return cached;
    } catch (_) {}

    if (sp.fileUrl) {
      try {
        const item = await graph(`/shares/${shareId(sp.fileUrl)}/driveItem?$select=id,name,parentReference`, accessToken);
        if (item && item.id && item.parentReference) {
          return guardarUbicacion({ driveId: item.parentReference.driveId, itemId: item.id });
        }
      } catch (err) {
        if (err.status === 401) throw err;
        // Se prueba con la búsqueda por nombre
      }
    }

    if (!sp.hostname || !sp.sitePath || !sp.fileName) {
      throw new ErrorConexion("Falta configurar la ubicación del Excel en config.js.", { tipo: "config" });
    }
    const site = await graph(`/sites/${sp.hostname}:${sp.sitePath}?$select=id,displayName`, accessToken);
    const drives = await graph(`/sites/${site.id}/drives?$select=id,name`, accessToken);
    const wanted = sp.fileName.toLowerCase();
    const q = encodeURIComponent(sp.fileName.replace(/'/g, "''"));
    for (const drive of drives.value || []) {
      const found = await graph(`/drives/${drive.id}/root/search(q='${q}')?$select=id,name,parentReference`, accessToken);
      const hit = (found.value || []).find((i) => (i.name || "").toLowerCase() === wanted);
      if (hit) return guardarUbicacion({ driveId: drive.id, itemId: hit.id });
    }
    throw new ErrorConexion(`No se encontró "${sp.fileName}" en el sitio ${sp.sitePath}.`, {
      tipo: "no-encontrado",
      detalle: "Revisá el nombre del archivo y del sitio en config.js, y que tu usuario tenga acceso.",
    });
  }

  /*
   * Descarga el Excel. Se usa la propiedad @microsoft.graph.downloadUrl (URL
   * temporal ya autorizada) porque el endpoint /content responde con una
   * redirección que el navegador no puede seguir con el encabezado Authorization.
   */
  async function leerExcel() {
    const accessToken = await token();
    let loc;
    try {
      loc = await ubicarArchivo(accessToken);
    } catch (err) {
      throw traducir(err);
    }
    let item;
    try {
      item = await graph(`/drives/${loc.driveId}/items/${loc.itemId}`, accessToken);
    } catch (err) {
      try {
        sessionStorage.removeItem(LOCATION_KEY);
      } catch (_) {}
      throw traducir(err);
    }
    const url = item["@microsoft.graph.downloadUrl"];
    if (!url) throw new ErrorConexion("SharePoint no devolvió el enlace de descarga del Excel.", { tipo: "graph" });
    const res = await fetch(url);
    if (!res.ok) {
      throw new ErrorConexion("No se pudo descargar el Excel desde SharePoint.", {
        tipo: "descarga",
        detalle: res.status + " " + res.statusText,
      });
    }
    return {
      buffer: await res.arrayBuffer(),
      meta: {
        nombre: item.name,
        webUrl: item.webUrl,
        modificado: item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime) : null,
        modificadoPor:
          (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || null,
      },
    };
  }

  function traducir(err) {
    if (!(err instanceof ErrorConexion)) {
      return new ErrorConexion("No se pudo conectar con SharePoint.", { tipo: "red", detalle: err.message });
    }
    if (err.status === 401 || err.status === 403) {
      return new ErrorConexion("Tu usuario no tiene acceso al Excel en SharePoint.", {
        tipo: "sin-acceso",
        detalle: err.detalle,
      });
    }
    if (err.status === 404) {
      return new ErrorConexion("No se encontró el Excel en SharePoint.", { tipo: "no-encontrado", detalle: err.detalle });
    }
    return err;
  }

  global.SharePoint = { estaConfigurado, iniciar, usuario, ingresar, salir, leerExcel, ErrorConexion };
})(window);
