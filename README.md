# Reglas de negocio consensuadas · Gobierno de Datos

Tablero web que muestra las reglas de negocio acordadas con los pilares (sedes, cambios de servicio y especialidad, unidades de negocio, exclusiones, Biogenómica y glosario) leyendo **en vivo** el Excel `DataGob_Reglas_Negocio.xlsx` del sitio de SharePoint **Qlik-Reportes**.

Se publica con **GitHub Pages** y exige iniciar sesión con la **cuenta de Microsoft de la empresa** (Entra ID / AD).

| Dato | Valor |
|---|---|
| Repositorio | `https://github.com/BautistaPalatDM/DATA-GOV` |
| Dirección del tablero | `https://bautistapalatdm.github.io/DATA-GOV/` |
| URI de redirección para IT (paso 3) | `https://bautistapalatdm.github.io/DATA-GOV/redirect.html` |

## Cómo funciona

```
Persona ──► abre el link ──► "Ingresar con Microsoft" (cuenta de la empresa)
                                       │
                                       ▼
            el navegador lee el Excel en SharePoint con los permisos de esa persona
                                       │
                                       ▼
                         arma el tablero en el navegador
```

Tres capas protegen los datos:

1. **La página no contiene datos.** En GitHub hay solo código. Sin iniciar sesión, se ve únicamente el botón de ingreso.
2. **Solo entran cuentas de la empresa.** La app está registrada en el Entra ID de Diagnóstico Maipú como *inquilino único*. Opcionalmente, IT puede limitarla a un grupo de personas (paso 3.6).
3. **Mismos permisos que SharePoint.** Aun logueada, cada persona ve el tablero solo si puede leer el Excel.

Además, los cambios en el Excel se ven al recargar la página o con el botón *Actualizar*.

## Estructura

```
app/                         ← lo que se publica en GitHub Pages
  index.html                 página principal (incluye la política de seguridad de contenido)
  redirect.html              página de retorno del inicio de sesión (requerida por MSAL v5)
  config.js                  ★ configuración: tenant, app, ubicación del Excel, link al MAESTRO
  js/app.js                  interfaz del tablero
  js/modelo.js               cómo se interpreta cada hoja del Excel
  js/sharepoint.js           inicio de sesión y lectura desde SharePoint
  js/xlsx-reader.js          lector de .xlsx propio, sin dependencias
  vendor/                    MSAL.js 5.30.0 (librería oficial de Microsoft, licencia MIT)
  css/, fonts/               estilos y tipografías (licencia OFL)
.github/workflows/publicar.yml  publica app/ en GitHub Pages en cada cambio de main
```

---

## Puesta en marcha (una sola vez)

### Paso 1 · Crear el repositorio y subir el código

1. En GitHub: **New repository** (en este caso ya existe: `BautistaPalatDM/DATA-GOV`).
   - Con **GitHub Free**, Pages solo publica desde repositorios **públicos**.
   - Con GitHub Pro, Team o Enterprise se puede usar un repositorio **privado**.
   - En cualquier caso, la página publicada es accesible por link. Solo con GitHub Enterprise Cloud se puede publicar de forma privada.
   - Los archivos del sitio, incluido `config.js`, no tienen datos ni secretos: solo identificadores y la dirección del sitio de SharePoint.
2. Subir el contenido de esta carpeta, incluida la carpeta oculta `.github`.
   - **Desde la web:** *uploading an existing file* y arrastrar todo.
   - **Con git:**
     ```bash
     git init && git add . && git commit -m "Tablero de reglas de negocio"
     git branch -M main
     git remote add origin https://github.com/BautistaPalatDM/DATA-GOV.git
     git push -u origin main
     ```

### Paso 2 · Activar GitHub Pages

1. En el repo: **Settings** → **Pages** → **Build and deployment** → **Source:** elegir **GitHub Actions**.
2. Ir a **Actions** → **Publicar en GitHub Pages** → **Run workflow** (o hacer cualquier cambio en `main`).
3. Al terminar, el sitio queda en:
   `https://bautistapalatdm.github.io/DATA-GOV/`
   Todavía va a mostrar "Falta conectar con Microsoft 365". Es normal hasta completar los pasos 3 y 4.

### Paso 3 · IT: registrar la aplicación en Microsoft Entra

1. Entrar a **Microsoft Entra admin center** (entra.microsoft.com) → **Registros de aplicaciones** → **Nuevo registro**.
2. Completar:
   - **Nombre:** `Reglas de Negocio DM`
   - **Tipos de cuenta admitidos:** *Solo cuentas de este directorio organizativo (inquilino único)*
   - **URI de redirección:** plataforma **Aplicación de página única (SPA)** con
     `https://bautistapalatdm.github.io/DATA-GOV/redirect.html`
     Tiene que ser exacta, con la carpeta del repo y `redirect.html`. Para pruebas locales se puede agregar también `http://localhost:8080/redirect.html`.
3. En **Permisos de API** → **Agregar un permiso** → **Microsoft Graph** → **Permisos delegados**, agregar `Files.Read.All` y `Sites.Read.All`.
4. **Conceder consentimiento de administrador** para la organización.
5. Copiar de **Información general** el **Id. de directorio (inquilino)** y el **Id. de aplicación (cliente)**.
6. **Opcional, para que entren solo ciertas personas:** en **Aplicaciones empresariales** → `Reglas de Negocio DM` → **Propiedades** → **¿Se requiere asignación?** = **Sí** → **Guardar**. Después, en **Usuarios y grupos**, agregar el grupo del negocio. Quien no esté asignado no puede iniciar sesión, aunque tenga cuenta de la empresa.

> **Para seguridad informática:** los permisos son *delegados*. La app nunca lee más de lo que puede leer la persona que inició sesión, y solo lee. No hay secreto de cliente ni servidor intermedio. El token queda en `sessionStorage` y se borra al cerrar la pestaña.

### Paso 4 · Completar `config.js`

1. En GitHub, abrir `app/config.js` → ícono del lápiz.
2. Reemplazar:
   - `tenantId` → Id. de directorio (inquilino) del paso 3
   - `clientId` → Id. de aplicación (cliente) del paso 3
   - `maestroUrl` → link de SharePoint del *MAESTRO - GOBIERNO DE DATOS*
3. **Commit changes**. La publicación es automática y tarda uno o dos minutos; se ve en la pestaña **Actions**.

### Paso 5 · Dar acceso al negocio

- Cada persona necesita **permiso de lectura** sobre `DataGob_Reglas_Negocio.xlsx` en Qlik-Reportes. Lo más práctico es un grupo de lectura en el sitio.
- Si se activó el paso 3.6, la persona también tiene que estar asignada a la app (conviene usar el mismo grupo).
- Compartir el link del paso 2.

### Paso 6 (opcional) · Fijar la ubicación del Excel

La primera vez, la app ubica el archivo por su link y, si no lo encuentra, por su nombre dentro del sitio. Para que cargue más rápido:

1. Abrir el tablero → **F12** → **Consola**.
2. Copiar los valores `driveId` e `itemId` que aparecen.
3. Pegarlos en `app/config.js` → `sharepoint.driveId` y `sharepoint.itemId`.

Si el Excel se reemplaza por otro archivo, vaciar esos dos valores y actualizar `fileUrl` con el link nuevo.

---

## Si cambia la estructura del Excel

- **Columnas:** se buscan por el **nombre del encabezado** (fila 1), no por posición. Se pueden reordenar o agregar columnas.
- **Hoja renombrada:** actualizar `hojas` en `config.js`.
- **Columna o hoja faltante:** el tablero muestra un aviso amarillo, *Revisar el Excel*, que dice qué falta.

| Hoja | Columnas que usa el tablero |
|---|---|
| `SEDES` | NUMERO, ORIGEN_ID, NOMBRE_SEDE, NOMBRE_SEDE_PARSEADO, TIPO_SEDE, EMPRESA, MODALIDAD, MODELO_NEGOCIO, FUENTE, CODIGO_ORIGEN, OBSERVACIONES_GOBIERNO |
| `ESPECIALIDAD cambia SERVICIO` | ACTIVA, FUENTE, ID_SERVICIO_FUENTE, SERVICIO_FUENTE, ID_ESPECIALIDAD_FUENTE, ESPECIALIDAD_FUENTE, SERVICIO_nuevo, ID_SERVICIO_nuevo, ESPECIALIDAD_nueva, ID_ESPECIALIDAD_nueva |
| `UNIDAD NEGOCIO` | ACTIVA, FUENTE, UNIDAD_NEGOCIO_actual, Servicio actual, ID_SERVICIO_FUENTE, Especialidad actual, ID_ESPECIALIDAD_FUENTE, UNIDAD_NEGOCIO_nueva |
| `ESTUDIO cambia SERV-ESPEC` | ACTIVA, FUENTE, PRA_ID, PRA_CODIGO, PRA_REGION, PRA_DESCRIPCION, ESPECIALIDAD_FUENTE, ESPECIALIDAD_ID_FUENTE, ID_ESPECIALIDAD_nueva, ESPECIALIDAD_nueva, ID_SERVICIO_nuevo, SERVICIO_nuevo |
| `EXCLUIR` | ACTIVA, FUENTE, ID_SERVICIO_FUENTE, SERVICIO_FUENTE, ID_ESPECIALIDAD_FUENTE, ESPECIALIDAD_FUENTE, ID_ESTUDIO_FUENTE, NOMBRE_ESTUDIO |
| `PRESTACIONES BIOGENOMICA` | COD SGM, COD DM, COD LM, COD CLIPPER, Nombre del Estudio, ESPECIALIDAD, SERVICIO |
| `NUEVA NORMALIZACION` | Celda `FUENTE` como encabezado de procesos. A su izquierda, la descripción y la marca `*`; a la derecha, una columna por proceso. |

Reglas que aplica el tablero:

- **Sede completa:** tiene definidos los cinco datos (sede normalizada, tipo, empresa, modalidad y modelo de negocio). Un dato vacío, `-` o con el texto "A DEFINIR" cuenta como *Falta definir*.
- **Regla activa:** `ACTIVA = SI`.
- **Alcance de una exclusión:** si tiene `ID_ESTUDIO_FUENTE`, se excluye el estudio; si tiene `ID_ESPECIALIDAD_FUENTE`, la combinación servicio + especialidad; si solo tiene `ID_SERVICIO_FUENTE`, todo el servicio.
- **Término pendiente de revisión:** el glosario tiene un `?` en la descripción o en algún proceso.

## Probar en una computadora

```bash
cd app
python -m http.server 8080
```

Abrir `http://localhost:8080`:

- **Sin `tenantId`/`clientId`:** ofrece abrir una copia local del Excel, que se lee en el navegador y no se sube a ningún lado.
- **Con la configuración completa:** inicia sesión con Microsoft. Requiere que `http://localhost:8080/redirect.html` esté registrada en el paso 3.
- **Agregando `?archivo` a la URL:** la pantalla de ingreso ofrece también abrir una copia local. Esto funciona solo en `localhost`, nunca en el sitio publicado.

## Solución de problemas

| Lo que se ve | Causa probable | Qué hacer |
|---|---|---|
| "Falta conectar con Microsoft 365" | `tenantId`/`clientId` vacíos o mal copiados | Paso 4 |
| `AADSTS50011` al iniciar sesión | La URI de redirección no coincide | Registrar exactamente `https://bautistapalatdm.github.io/DATA-GOV/redirect.html` (paso 3.2) |
| `AADSTS65001` o pedido de aprobación | Falta el consentimiento de administrador | Paso 3.4 |
| `AADSTS50105` | La persona no está asignada a la app | Agregarla al grupo asignado (paso 3.6) |
| "No tenés acceso al Excel en SharePoint" | Sin permiso de lectura sobre el Excel | Paso 5 |
| "No se encontró el Excel en SharePoint" | El archivo se reemplazó, movió o renombró el sitio | Actualizar `sharepoint` en `config.js` y vaciar `driveId`/`itemId` |
| Aviso amarillo "Revisar el Excel" | Una hoja o columna cambió de nombre | Ajustar el Excel o `hojas` en `config.js` |
| El sitio no se actualiza | Falló la publicación | Ver la pestaña **Actions** y que Pages esté en *GitHub Actions* |

## Notas para IT

- **Dependencias externas en tiempo de ejecución:** ninguna. MSAL.js y las fuentes están incluidas en el repo. No se usan CDNs.
- **Conexiones que hace el navegador:** `login.microsoftonline.com` (inicio de sesión), `graph.microsoft.com` (ubicar el archivo) y `*.sharepoint.com` (descarga del Excel con un enlace temporal que entrega Graph). La política de seguridad de contenido, en una etiqueta `<meta>` de `index.html` y `redirect.html`, permite solo esos destinos.
- **Almacenamiento:** el token queda en `sessionStorage` y se borra al cerrar la pestaña. El Excel se procesa en memoria.
- **Repositorio:** no debe contener copias del Excel. `.gitignore` excluye `.xlsx`, `.xls` y `.csv`.
