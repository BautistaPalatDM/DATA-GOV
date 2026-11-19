// Devuelve la respuesta del inicio de sesión a la página principal (MSAL v5).
window.msalRedirectBridge.broadcastResponseToMainFrame().catch(function (err) {
  console.error("[Reglas de negocio] Error al volver del inicio de sesión:", err);
  window.location.replace("./");
});
