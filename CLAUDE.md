# Contexto del proyecto: refriceiba-sync

Este archivo resume todo lo hecho hasta el 17/11/2026 (incluye el modo movimientos) para continuar en Claude Code. Léelo antes de cambiar código. El `README.md` describe la herramienta para usuarios; aquí están las decisiones, los datos reales y lo pendiente.

## El negocio

- **Refrigeraciones La Ceiba** (refriceiba.com), tienda de repuestos de refrigeración en Cabudare, Lara (Venezuela). La lleva el hermano de Andrés Lugo; Andrés hace la parte técnica.
- Tres sistemas:
  - **Valery**: software de facturación venezolano. **Es la fuente de verdad del inventario.** Toda venta (tienda, web y Cashea) se factura ahí.
  - **Cashea**: marketplace de compra en cuotas. Portal de comercios: `merchants.cashea.app`. **No tiene API pública ni documentación**; todo se hace subiendo planillas Excel en el portal.
  - **WordPress + WooCommerce** (refriceiba.com): tienda web.
- Objetivo: que el hermano haga el **cierre del día** cargando el reporte de Valery y que el stock de Cashea (y luego WordPress) quede igual al de Valery.

## Qué es este repo

Página estática (sin build) para GitHub Pages: `https://wanderlust51.github.io/refriceiba-sync/` (Pages se activa en Settings → Pages → rama `main`, carpeta raíz; desde esta sesión no se pudo activar por permisos). Repo público `WanderLust51/refriceiba-sync`, rama `main`. Todo corre en el navegador; los archivos del negocio nunca salen de la computadora.

```
index.html          Interfaz (CSS inline)
js/sync.js          Lógica pura; funciona en navegador y en Node (UMD → window.RefriSync)
js/app.js           Interfaz: carga de archivos, alertas, tabla, descargas, localStorage
vendor/             SheetJS 0.18.5 (lee los .xls de Valery), JSZip 3.10, pdf.js 3.11.174 (facturas PDF;
                    siempre con isEvalSupported: false por la vulnerabilidad CVE-2024-4367 de esa versión)
test/verificar.js   node test/verificar.js VALERY.xls PLANILLA.xlsx [SALIDA.xlsx]
```

Pruebas en Node: `npm install --no-save xlsx@0.18.5 jszip@3 pdfjs-dist@3.11.174` (en Node usar `pdfjs-dist/legacy/build/pdf.js`). **Nunca subir al repo archivos .xls/.xlsx/.csv del negocio**: el repo es público y `.gitignore` ya los excluye. Para probar, pedirle a Andrés los archivos y trabajar fuera del repo.

## Archivos de entrada (formatos reales)

### Valery: "Listado de Productos"
- Cómo se obtiene: **Reportes → Listados → Listado de productos → Exportar**. Formato `.xls` viejo (BIFF), una hoja llamada "Listado de Productos de Valery ".
- Encabezado en la fila 1. Columnas: `Código`, `Nombre`, `Dpto. Descrip.`, `Und. Ppal.`, `Existencia Actual`, `Precio Máximo`, etc. 576 productos.
- `Código` es numérico (1, 2 … 577), sin prefijo.
- **Trampa:** la última fila es de totales; en la columna Código trae el conteo de productos (576), **que choca con un código real** (576 = caja de breakers). Se descarta porque no tiene `Nombre`.
- Existencias con decimales para productos por kg o metros (refrigerantes, tubería de cobre) y alguna negativa (cinta PVC en −1).
- Los precios de Valery no son confiables: mezclan USD y Bs en la misma columna. **Solo se sincroniza stock, nunca precios.**
- **El .xls no trae fecha adentro.** Se usa `file.lastModified` (hora de exportación).

### Cashea: "Planilla de precio, cuotas y stock"
- Cómo se obtiene: **Inventario → Actualizar productos → Planilla de precio, cuotas y stock: Descargar → seleccionar todas las categorías → Descargar planilla.** Se sube en la misma pantalla con **Subir**.
- Una sola hoja. Fila 2 = encabezados: A `Título`, B `SKU`, C `Stock`, D cuotas, E precio. Datos desde la fila 5. ~442 publicaciones.
- Reglas de Cashea (de su propia pantalla): no modificar columnas, filas, SKU ni títulos; **cada planilla descargada se puede subir una sola vez**; descargar una nueva cada vez. Cashea parece identificar productos por fila, así que **nunca se borran ni reordenan filas** (Andrés preguntó por un archivo de una sola fila y se descartó por esto).
- **Trae fecha adentro:** las entradas del zip tienen la hora de generación (UTC); JSZip la expone en `file.date`. Se usa esa.
- SKU en Cashea: `REF-123` (mayúscula) para la mayoría; 178 todavía son solo el número (`123`) mientras los migran a `REF-`; 19 combos con SKU propio (`13GMCC`, `COMB-MM`…).
- Hay otra opción, "Planilla de catálogo completo" (todos los campos), y una plantilla de **carga masiva** ("Cargar_Productos", una hoja por categoría, 20 columnas, validaciones de lista en filas 6–104 ≈ 99 productos por hoja) para crear productos nuevos.

### Valery: "Relación de Ventas por Producto" (modo movimientos)
- Cómo se obtiene: **Reportes → Ventas → Relación de ventas → Relación de ventas por producto → Exportar** (.xls).
- Columnas: `Código`, `Nombre`, `Departamento Código`, `Departamento Nombre`, `Cant (Und. Principal)`, totales en Bs y moneda contable. Solo trae los productos vendidos. Última fila = totales sin nombre.
- No trae fecha ni canal (no distingue ventas de Cashea).

### Valery: factura de compra en PDF (modo movimientos)
- Cómo se obtiene: **Reimprimir → Compras → Facturas de compras** ("Recepción de factura de compra"), en PDF (ReportBuilder).
- Encabezado: proveedor, RIF (`J-50161271-4`), número (`2736`) bajo "FACTURA DE COMPRA", "Fecha Emisión" (`16/nov./2026`).
- Columnas: Código, Nombre, Cantidad Ord., Bonif., Costo Lista, % Dcto 1, % Dcto 2, Sub Total, Impuesto, Total. Números con formato `1.234,56`.
- El texto extraído línea por línea se monta; se lee por posición (ver `parseCompra`). Validado con la factura 2736 de COLDWA: 15 líneas, todas cuadran y la suma da el Total Lineas (269.910,56).
- Esa factura trae códigos nuevos (640–647) que no estaban en el listado del 13/11.

### WordPress / WooCommerce
- Exportación nativa (Productos → Exportar) en CSV con encabezados en español: `ID`, `Tipo`, `SKU`, `Nombre`, `Publicado`, `¿Existencias?`, `Inventario`, `Precio normal` (coma decimal, ej. `6,5`), `Categorías`, `Imágenes`, `Marcas`…
- 595 productos, todos simples. SKU `REF-123`, salvo `573` y `576` sin prefijo (hay que corregirlos en WP).
- 2 productos sin inventario gestionado: espuma COVO (REF-530) y bomba de vacío REF-552 (en "reserva").
- El plugin que tenían para exportar a Cashea (`cashea-export`) no cumple la plantilla (categoría equivocada, faltan marca/peso/condición, marca comercial con texto distinto). No se usa.

## Reglas de sincronización (implementadas en js/sync.js)

| Caso | Qué hace |
|---|---|
| SKU `REF-123` o `123` | Cruza con el código 123 de Valery (`/^(?:REF-)?0*(\d+)$/i`) |
| Decimales | Parte entera, redondeando antes a 6 decimales (9,5 → 9; 2,9999999 → 3) |
| Negativo | 0 |
| SKU sin número (combos) | No se toca |
| Código que no existe en Valery | No se toca |
| Título de Cashea sin ninguna palabra en común con el nombre en Valery | No se toca y se alerta (`sameProduct`: palabras de 3+ letras o números; prefijo de 4+ letras cuenta, "pega" = "pegamento") |
| Código publicado dos veces en Cashea | Ambas reciben el stock; se alerta para pausar una |
| Más de 50 cambios | Pide confirmación antes de descargar |
| Planilla con más de 60 min | Alerta |
| Reporte de Valery que no es de hoy | Alerta |

**Edición de la planilla:** se modifica solo el `<v>` de las celdas `C{fila}` directamente en el XML de la hoja (regex sobre `<c r="C5" ...>`), quitando el atributo `t` si lo hubiera. El resto del zip queda igual. Verificado: cero celdas distintas fuera de la columna C, validaciones intactas, y resultado idéntico al del script Python original.

## Dos modos de cierre (selector en el paso 1)

- **Inventario completo:** Listado de productos + planilla de Cashea. Compara todos los productos.
- **Movimientos del día:** Relación de ventas (.xls) y/o facturas de compra (PDF, varias) + planilla de Cashea.
  `existencia nueva = existencia de Valery guardada en el último cierre − ventas + compras`, y solo se tocan en Cashea los productos con movimiento (los demás quedan "Sin movimiento").
  - Se parte de la existencia guardada, **no** del stock de Cashea: la planilla ya trae descontadas las ventas de Cashea y la Relación de ventas también las incluye (todo se factura en Valery), así que se descontarían dos veces.
  - Requiere un cierre guardado en esa computadora (si no, pide un inventario completo primero). Sugiere un inventario completo si el último tiene más de 7 días.
  - Huella por documento (`compra:RIF:número`, `ventas:sha256`) en `refriceiba-sync/documentos`; un documento ya aplicado no se aplica de nuevo. Si los documentos son justo los del último cierre, se **regenera** la planilla sin reaplicar (`replayMovements`), útil si falló la subida.
  - Líneas de compra que no cuadran (cantidad × costo × descuentos ≠ subtotal) no se aplican. Unidades recibidas = Ord. + Bonif.
  - Avisos: productos nuevos sin existencia guardada, negativos, con movimiento pero no publicados en Cashea, cierre previo el mismo día (ventas contadas dos veces), Cashea vendió más de lo reportado.
  - Al descargar se guarda el cierre (modo `movimientos`, con `docs` y `ultimoCompleto`) y se marcan los documentos.
- El modo elegido se recuerda en `refriceiba-sync/modo`.
- Probado el 17/11 con la Relación de ventas (8 productos, 14,6 unid.) y la factura 2736 (15 productos, 65 unid.): 12 cambios en Cashea, 2 productos vendidos sin publicar (274 gas R22, 323 capacitor 2,5UF), 8 productos nuevos y 1 negativo (566). Solo cambian celdas de la columna C.

## Ventas del día por canal (localStorage)

Clave `refriceiba-sync/cierres`: hasta 14 entradas, una por día (gana la última del día):
`{ t: fecha del reporte de Valery, guardado, archivo, valery: {código: existencia}, cashea: {SKU: stock subido}, casheaT }` (~12–17 KB).

- Se guarda al descargar la planilla para Cashea (o al cargar, si hay 0 cambios).
- Base de comparación: el último cierre de un **día anterior** al del reporte de Valery actual. Varios cierres el mismo día siguen comparando contra el día anterior.
- Vendido por Cashea = stock subido a Cashea ayer − stock que trae hoy la planilla.
- Vendido tienda/web = lo que bajó Valery − lo vendido por Cashea.
- Entradas = lo que subió Valery.
- Alertas: "¿falta facturar?" (Cashea vendió más de lo que bajó Valery) y "subido a mano en Cashea" (el stock de Cashea subió).
- Supone que Cashea descuenta su stock al vender (**sin confirmar**) y que la planilla de ayer se subió.
- Si no hay historial (otra computadora, datos borrados), se puede cargar a mano un reporte de Valery anterior: solo muestra el movimiento neto, sin separar por canal.

## Historial y hallazgos

- **13/11:** prueba con un solo producto (REF-119, Control Remoto Universal 1000 Códigos, 10 → 6) subiendo la planilla completa con una sola celda cambiada. **Funcionó.**
- **13/11:** sincronización inicial con ~100 cambios. Salió bien según Andrés.
- Diferencias iniciales grandes: 154 productos con más stock en la web que en Valery. Los stocks web se habían cargado a mano.
- **SKU equivocados en Cashea** (la página los excluye):
  - `252` "Toma 220V Tipo China P/Empo" → en Valery es Tapa 270 Beige; la toma es el **261**.
  - `262` "Tapa Beige 240" → en Valery es un enchufe trifásico; probablemente **252 o 259** (por confirmar).
  - `510` "Lona aires 9–18 KBTU" → en Valery es tubería de cobre 5/8; la lona **no existe en Valery**.
  Andrés dijo que están corrigiendo esto.
- **Duplicados en Cashea:** `REF-242`/`242` (adaptador válvula R410A) y `REF-158`/`158` (lona 18–36 mil BTU). Hay que pausar uno de cada par.
- **155 productos de Valery/WordPress no están en Cashea.** Se generó `completar_para_cashea.xlsx` (marca y peso a completar; WP no tiene peso en ningún producto y solo 13 tienen marca) y un generador de carga masiva (`generar_carga_cashea.py`, fuera de este repo). 28 productos (Automotriz, Línea Blanca, Sin categorizar) necesitan una plantilla con esas categorías. **Pendiente** de que el hermano complete la hoja.

## Problema abierto (prioridad)

**16/11:** el viernes 13 por la tarde, después de la sincronización inicial, el hermano hizo otro cierre: había vendido ~10 cosas por Valery y la página solo mostró 3 cambios para Cashea. Causas posibles, sin confirmar:
1. Productos vendidos que no están publicados en Cashea (155 de 576).
2. Varias unidades del mismo producto (la página cuenta productos, no unidades).
3. Ventas hechas por Cashea: su stock ya había bajado solo.
4. Productos por kg/m donde no cambia la parte entera.
5. Productos excluidos (combos, SKU equivocados, ya en 0).
6. Reporte de Valery exportado antes de terminar de facturar, o ventas registradas como documentos que no descuentan inventario (pedido, presupuesto, nota de entrega, en espera), o reporte filtrado por depósito.

Ese día no se mostró el panel de ventas porque la base y el cierre fueron el mismo día. **Pendiente:** el hermano va a pasar los archivos de ese cierre y la lista de ventas para cruzarlos. Se sugirió una prueba en Valery: comparar la ficha del producto con el listado exportado justo después de una venta. No hay documentación pública de Valery.

## Pendientes

1. Diagnosticar el problema del 13/11 con los archivos reales. (El modo movimientos ya muestra qué productos con movimiento no están publicados en Cashea; es la causa más probable.)
1b. Probar el modo movimientos en un cierre real: confirmar que el rango de fechas de la Relación de ventas coincide con el período desde el último cierre, y que la cantidad de la factura de compra está en la unidad principal.
2. **WordPress:** WooCommerce trae un importador CSV (Productos → Importar → "Actualizar productos existentes"). Hace coincidir por ID o SKU y **omite las filas sin coincidencia**, así que no crea productos. Plan: que la página genere un CSV `SKU,Inventario` desde Valery (todos los códigos como `REF-{código}`), excluyendo productos sin inventario gestionado. Primero, una prueba con un solo producto, como se hizo con Cashea. Requiere rol Gestor de tienda o Administrador. Descartado: llamar la API de WooCommerce desde la página, porque expondría las claves en un repo público.
3. Alerta de "productos nuevos en Valery que no están en Cashea": hecha en modo movimientos y para los productos con movimiento en modo completo.
4. Confirmar que Cashea descuenta su stock al vender (necesario para las ventas por canal).
5. Corregir en WordPress los SKU `573` y `576` a `REF-`.
6. Publicar en Cashea los 155 faltantes cuando esté completa la hoja de marca/peso.

## Cómo trabajar con Andrés

- Escribe en español. Prefiere ir paso a paso y probar con un solo producto antes de cambios masivos.
- Antes de subir algo que cambie datos reales, probar con los archivos reales y verificar que no cambie nada fuera de lo esperado.
- La página es para el hermano, que no es técnico: nada de terminal, mensajes claros, alertas que expliquen qué hacer.
- Commits con autor Andrés Lugo. El push a `main` publica directo en GitHub Pages.
