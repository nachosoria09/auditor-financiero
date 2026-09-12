# Manual de Usuario — Auditor Financiero Personal

Esta app registra tus ingresos, gastos, préstamos y ahorros del mes, y genera un cierre mensual con crítica honesta incluida. Tus datos se guardan en **tu propia Google Sheet** — nadie más los ve, ni siquiera quien te pasó el link de la app.

Una ventaja extra: los datos se guardan en la planilla como **filas normales y legibles** (fecha, tipo, categoría, monto, nota) — si en algún momento querés mirarlos directamente en Google Sheets, sin abrir la app, vas a poder entenderlos sin problema.

Cada persona que use esta app necesita hacer esta configuración **una sola vez**.

---

## Antes de empezar: palabras que vas a ver

No hace falta entender esto a fondo, pero para que no te frene ninguna palabra rara en el camino:

- **Enlace / link / URL** — son la misma cosa: un texto que arranca con `https://` y que, si lo pegás en un navegador, te lleva a algún lado. En este manual vas a copiar y pegar varios.
- **Google Apps Script** — es una herramienta que ya viene adentro de Google Sheets (no hay que instalar nada aparte). Sirve para que la app y tu planilla se puedan "hablar" entre sí. Vos no vas a programar nada: solo vas a copiar un texto que ya está armado (Paso 2) y pegarlo.
- **"Implementar" / "Publicar"** — es el botón que le dice a Google "dejá que esta herramienta funcione y dame un enlace para usarla". Es un solo click, Google hace el resto.
- **"Aplicación web"** — una opción que hay que elegir en un menú desplegable durante el Paso 3. No implica que tengas que crear ninguna página ni saber diseño.

---

## Paso 1 — Crear tu Google Sheet

1. Andá a [sheets.google.com](https://sheets.google.com) y creá una planilla nueva.
2. Ponele el nombre que quieras (ej: "Mi Auditor Financiero").
3. No hace falta que le crees hojas ni columnas — la app las crea sola la primera vez que la uses.

---

## Paso 2 — Agregar el código que conecta la planilla con la app

1. Dentro de la planilla: **Extensiones → Apps Script**.
2. Borrá el código de ejemplo que aparece y pegá exactamente este:

```javascript
// === Backend Apps Script para Auditor Financiero (v2 — filas legibles) ===
// Pegar este código en Extensiones > Apps Script de tu Google Sheet.
// Crea automáticamente 3 hojas: Movimientos, Prestamos, Categorias.
// Cada fila es un dato real y legible — podés abrir la Sheet y entenderla
// sin necesidad de la app.

function hojaMovimientos_() {
  return obtenerOCrearHoja_('Movimientos', ['ID','Usuario','Fecha','Tipo','Categoria','Monto','Nota']);
}
function hojaPrestamos_() {
  return obtenerOCrearHoja_('Prestamos', ['Usuario','Nombre','Cuota','Actual','Total']);
}
function hojaCategorias_() {
  return obtenerOCrearHoja_('Categorias', ['Usuario','Tipo','Nombre']);
}

function obtenerOCrearHoja_(nombre, encabezados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(encabezados);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Convierte todas las filas de una hoja en objetos {columna: valor},
// arreglando el caso en que Google Sheets convirtió una fecha en tipo Date.
function filasComoObjetos_(sheet, encabezados) {
  var datos = sheet.getDataRange().getValues();
  var salida = [];
  for (var i = 1; i < datos.length; i++) {
    var obj = {};
    for (var c = 0; c < encabezados.length; c++) {
      var val = datos[i][c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[encabezados[c]] = val;
    }
    salida.push(obj);
  }
  return salida;
}

function borrarFilasDeUsuario_(sheet, usuario, columnaUsuarioIdx) {
  var datos = sheet.getDataRange().getValues();
  for (var i = datos.length - 1; i >= 1; i--) {
    if (datos[i][columnaUsuarioIdx] === usuario) {
      sheet.deleteRow(i + 1);
    }
  }
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var accion = e.parameter.accion;
  var usuario = e.parameter.usuario;

  if (accion === 'cargarTodo') {
    var movimientos = filasComoObjetos_(hojaMovimientos_(), ['ID','Usuario','Fecha','Tipo','Categoria','Monto','Nota'])
      .filter(function (m) { return m.Usuario === usuario; });
    var prestamos = filasComoObjetos_(hojaPrestamos_(), ['Usuario','Nombre','Cuota','Actual','Total'])
      .filter(function (p) { return p.Usuario === usuario; });
    var categorias = filasComoObjetos_(hojaCategorias_(), ['Usuario','Tipo','Nombre'])
      .filter(function (c) { return c.Usuario === usuario; });
    return salida_({ ok: true, movimientos: movimientos, prestamos: prestamos, categorias: categorias });
  }

  return salida_({ ok: false, error: 'accion desconocida' });
}

function doPost(e) {
  var datos = JSON.parse(e.postData.contents);
  var accion = datos.accion;

  if (accion === 'agregarMovimiento') {
    hojaMovimientos_().appendRow([
      datos.id, datos.usuario, datos.fecha, datos.tipo, datos.categoria, datos.monto, datos.nota || ''
    ]);
    return salida_({ ok: true });
  }

  if (accion === 'eliminarMovimiento') {
    var sh = hojaMovimientos_();
    var filas = sh.getDataRange().getValues();
    for (var i = filas.length - 1; i >= 1; i--) {
      if (filas[i][0] === datos.id) { sh.deleteRow(i + 1); break; }
    }
    return salida_({ ok: true });
  }

  if (accion === 'guardarPrestamos') {
    var shP = hojaPrestamos_();
    borrarFilasDeUsuario_(shP, datos.usuario, 0);
    (datos.prestamos || []).forEach(function (p) {
      shP.appendRow([datos.usuario, p.name, p.cuota, p.current, p.total]);
    });
    return salida_({ ok: true });
  }

  if (accion === 'guardarCategorias') {
    var shC = hojaCategorias_();
    borrarFilasDeUsuario_(shC, datos.usuario, 0);
    ['ingreso', 'gasto', 'ahorro'].forEach(function (tipo) {
      (datos.categorias[tipo] || []).forEach(function (nombre) {
        shC.appendRow([datos.usuario, tipo, nombre]);
      });
    });
    return salida_({ ok: true });
  }

  return salida_({ ok: false, error: 'accion desconocida' });
}
```

3. Guardá el proyecto (ícono de disquete o `Ctrl+S`).

---

## Paso 3 — Publicar el script y obtener tu enlace

1. Arriba a la derecha, botón **Implementar → Nueva implementación** (es el botón que "activa" el código que pegaste).
2. En "Tipo", elegí **Aplicación web** (es la única opción que nos sirve, no hay que dudar entre varias).
3. Configurá estas dos cosas — son menús desplegables, tocás y elegís:
   - **Ejecutar como:** vos (tu cuenta de Google).
   - **Quién tiene acceso:** *Cualquier usuario* (esto es necesario para que la app pueda guardar datos sin pedirte iniciar sesión cada vez).
4. Hacé clic en **Implementar**.
5. Te va a aparecer una ventana para elegir tu cuenta de Google — elegí la misma con la que creaste la planilla.
6. **Vas a ver una pantalla que dice algo como "Google no verificó esta app" o "Esta app no está verificada".** Es esperable y no es un error: aparece porque el script lo creaste vos recién, no porque haya un problema de seguridad. Para continuar:
   - Hacé clic en **Avanzado** (o "Advanced", puede aparecer en inglés).
   - Después hacé clic en **Ir a [nombre del proyecto] (no seguro)** / **Go to (unsafe)**.
   - Por último, **Permitir** / **Allow**.
7. Al final te va a aparecer un cuadro con un texto largo que empieza con `https://` y termina en `/exec`. **Ese es tu enlace** — copialo tal cual, con el botón de copiar que aparece al lado.

Guardalo, lo vas a necesitar en el paso 4.

---

## Paso 4 — Abrir tu Auditor Financiero

Abrí el link de la app que te pasaron. La primera vez te va a mostrar una pantalla llamada **"Antes de arrancar"**, con dos casillas:

- **¿Cómo te llamás?** — poné tu nombre (es solo para separar tus datos de los de otras personas que usen la misma app).
- **Pegá el enlace que copiaste en el Paso 3.**

Apretá **"Guardar y entrar"**. La app va a probar la conexión con tu planilla:
- Si algo está mal (enlace incompleto, o falta algún paso de la publicación), te va a avisar en rojo qué revisar.
- Si todo funciona, vas a ver un cartel con **tu link personal completo** y un botón **"Copiar link"**.

**Guardá ese link** (en tus notas, favoritos del navegador, o mandátelo a vos mismo por WhatsApp). Es la forma de volver directo a tus datos sin tener que llenar el formulario de nuevo.

> ⚠️ Si perdés ese link y volvés a entrar sin haberlo guardado, la app te va a volver a mostrar el formulario — no pasa nada, tus datos siguen intactos en tu planilla. Solo tenés que volver a poner tu nombre (igual que la primera vez) y el mismo enlace del Paso 3.

---

## Uso diario

### Registrar un movimiento
- Elegís el tipo: **Ingreso**, **Gasto** o **Ahorro** (alcancía).
- Elegís la categoría (ver más abajo cómo crear las tuyas).
- Cargás el monto en pesos uruguayos y, si querés, una nota.
- Click en **Agregar**. La fecha se pone sola (la del día en que lo cargás), no hay que escribirla.

Podés cambiar el mes que estás mirando con el selector de arriba — es instantáneo, no vuelve a pedirle nada a la planilla, porque tus movimientos ya están cargados en la app.

### Categorías propias
Cada usuario tiene sus propias categorías, no vienen fijas. En la tarjeta **"Mis categorías"**:
- Escribís el nombre (ej: "Gastos de trabajo", "Crédito Hipotecario", "Bebidas Alcohólicas").
- Elegís si es Ingreso, Gasto o Ahorro.
- Click en el botón "+".

Podés borrar cualquier categoría que ya no uses con el botón ✕.

*Ejemplo:* si trabajás y a veces pagás cosas de la empresa que después te reintegran, podés crear la categoría de gasto "Gastos de trabajo" y la de ingreso "Reintegro empresa" — así queda todo separado y prolijo.

### Préstamos fijos
En la tarjeta **"Préstamos fijos"**, cargás:
- Nombre (ej: Auto).
- Cuota en pesos.
- En qué cuota vas actualmente (si ya veías pagando antes de usar la app, poné el número real, ej: 10).
- Total de cuotas (ej: 24).

Cada vez que pagás una cuota, apretás **"Registrar pago"**: suma +1 al contador y lo carga automáticamente como gasto del mes. Cuando llega al total, el préstamo se marca como **PAGADO**.

### Auditoría / Cierre de mes
Apretando **"Generar Auditoría"** se arma un reporte con:
- Ingresos totales.
- Gastos desglosados por categoría.
- Estado de todos tus préstamos.
- Total acumulado en la alcancía.
- Porcentaje de tus ingresos gastado en tabaco y bebidas.
- Saldo disponible real del mes.
- Una evaluación directa y sin filtro de cómo te fue.

---

## Privacidad

- Tus datos viven únicamente en la Google Sheet que vos creaste, en tu propia cuenta de Google.
- Nadie más — ni la persona que te compartió la app, ni Anthropic (Claude) — tiene acceso a esa planilla.
- Si varias personas de tu familia usan la misma app, cada una con su propia planilla y su propio link, los datos de cada uno quedan completamente separados.
