/*
 * Configuración del tablero "Reglas de negocio consensuadas".
 *
 * tenantId y clientId los entrega IT al registrar la aplicación en Microsoft Entra
 * (ver README.md, paso 1). No son secretos: identifican la app, no dan acceso.
 */
window.APP_CONFIG = {
  // ---- Microsoft Entra (completar con lo que entregue IT) ----
  tenantId: "COMPLETAR-TENANT-ID",
  clientId: "COMPLETAR-CLIENT-ID",

  // ---- Ubicación del Excel en SharePoint ----
  sharepoint: {
    hostname: "diagnosticomaipu685.sharepoint.com",
    sitePath: "/sites/Qlik-Reportes",
    fileName: "DataGob_Reglas_Negocio.xlsx",
    fileUrl:
      "https://diagnosticomaipu685.sharepoint.com/:x:/r/sites/Qlik-Reportes/_layouts/15/Doc.aspx?sourcedoc=%7BC0A75555-7C53-414E-B3F0-22B5338AEBE1%7D&file=DataGob_Reglas_Negocio.xlsx&action=default&mobileredirect=true",
    // Opcional: la primera vez que alguien entra, la consola del navegador (F12)
    // muestra driveId e itemId. Copiarlos acá hace la carga más rápida.
    driveId: "",
    itemId: "",
  },

  // ---- Link del botón "Abrir MAESTRO - GOBIERNO DE DATOS" ----
  // Pegar acá el link de SharePoint del MAESTRO. Si queda vacío, el botón no se muestra.
  maestroUrl: "",

  // ---- Nombres de las hojas del Excel ----
  // Si alguien renombra una hoja, actualizar acá.
  hojas: {
    sedes: "SEDES",
    especialidadServicio: "ESPECIALIDAD cambia SERVICIO",
    unidadNegocio: "UNIDAD NEGOCIO",
    estudios: "ESTUDIO cambia SERV-ESPEC",
    biogenomica: "PRESTACIONES BIOGENOMICA",
    excluir: "EXCLUIR",
    glosario: "NUEVA NORMALIZACION",
  },
};
