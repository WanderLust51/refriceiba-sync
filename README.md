# refriceiba-sync

Página para el cierre del día de **Refrigeraciones La Ceiba**: toma el reporte de existencias de **Valery** y actualiza el stock en la planilla de **Cashea**.

Valery es la fuente de verdad. La página no modifica Valery ni se conecta a Cashea: genera el archivo que luego se sube a mano en el portal de Cashea.

Todo corre en el navegador. Los archivos no se suben a ningún servidor; la página es estática y se publica en GitHub Pages.

## Uso diario

1. Facturar en Valery todas las ventas del día (tienda, web y Cashea).
2. **Valery:** Reportes → Listados → Listado de productos → Exportar.
3. **Cashea:** Inventario → Actualizar productos → *Planilla de precio, cuotas y stock*: Descargar → seleccionar todas las categorías → Descargar planilla. No abrirla en Excel.
4. Abrir la página, cargar los dos archivos y revisar los cambios.
5. Descargar la planilla generada y subirla en Cashea, en la misma sección (**Subir**). Hacerlo el mismo día: lo vendido por Cashea del día siguiente se calcula contra esa planilla.

Cada planilla de Cashea se puede subir una sola vez. Si algo falla, se descarga otra y se repite.

## Reglas

| Caso | Qué hace |
|---|---|
| SKU `REF-123` o `123` en Cashea | Se cruza con el código `123` de Valery |
| Existencia con decimales (kg, metros) | Se publica la parte entera: 9,5 → 9 |
| Existencia negativa | Se publica 0 |
| SKU sin número (combos) | No se toca; su stock se lleva a mano |
| Código que no existe en Valery | No se toca |
| El título de Cashea no comparte ninguna palabra con el nombre en Valery | No se toca y se avisa: probablemente el SKU está mal |
| Mismo código publicado dos veces en Cashea | Ambas reciben el stock y se avisa para pausar una |
| Más de 50 cambios en un cierre | Pide confirmación antes de descargar |
| Planilla de Cashea con más de 60 minutos | Avisa que conviene descargarla de nuevo |

## Ventas del día por canal

Al descargar la planilla, la página guarda en el navegador (`localStorage`, unos 12 KB por día, máximo 14 días) dos listas de números:

- `valery`: `{ código: existencia }` del reporte de Valery.
- `cashea`: `{ SKU: stock }` con el que queda Cashea después de subir la planilla.

En el siguiente cierre compara contra el último cierre **de un día anterior** al del reporte de Valery que se está usando. Si el mismo día se hace el cierre varias veces, se guarda solo el último y la comparación sigue siendo contra el día anterior.

| Dato | Cálculo |
|---|---|
| Vendido por Cashea | stock subido a Cashea ayer − stock que trae hoy la planilla de Cashea |
| Vendido por tienda y web | lo que bajó Valery − lo vendido por Cashea |
| Entradas | lo que subió Valery |
| ¿Falta facturar? | Cashea vendió más de lo que bajó Valery |
| Ajuste manual en Cashea | el stock de Cashea subió desde la planilla de ayer |

Supone que Cashea descuenta su stock al vender y que la planilla de ayer se subió. Las ventas y entradas de un mismo producto en el mismo día se compensan (se ve el neto).

**Fechas:** la planilla de Cashea trae adentro la hora en que Cashea la generó (fecha de las entradas del `.xlsx`), y se usa esa. El `.xls` de Valery no trae fecha; se usa la hora en que se exportó el archivo.

Si se usa otra computadora o se borran los datos del navegador, se pierde el historial. En ese caso se puede cargar a mano el reporte de Valery anterior y se ve el movimiento total, sin separar por canal.

Las pestañas **Suben** y **Bajan** filtran los cambios que se van a aplicar en Cashea.

## Cómo se edita la planilla

Solo se modifica la columna **Stock** de la planilla: se editan esas celdas directamente en el XML del `.xlsx`, sin regenerar el archivo. Títulos, SKU, precios, cuotas y la validación de Cashea quedan intactos. Cashea identifica los productos por fila, por eso no se borran ni se reordenan filas.

## Estructura

```
index.html          Interfaz
js/sync.js          Lógica (lectura de Valery, cruce, edición de la planilla). Funciona también en Node.
js/app.js           Comportamiento de la página
vendor/             SheetJS 0.18.5 (lee el .xls de Valery) y JSZip 3.10 (edita la planilla de Cashea)
test/verificar.js   Corre la lógica en Node con archivos reales
```

## Verificar con archivos reales

```bash
npm install --no-save xlsx@0.18.5 jszip@3
node test/verificar.js Listado_de_Productos.xls planilla_cashea.xlsx salida.xlsx
```

Imprime el resumen y escribe la planilla resultante. **No subir al repositorio archivos de Valery ni de Cashea**: el repo es público. El `.gitignore` ya excluye `.xls`, `.xlsx` y `.csv`.

## Publicación

GitHub Pages, rama `main`, carpeta raíz. No hay paso de compilación.

## Pendiente

- WordPress: generar el CSV de `SKU` + `Inventario` para el importador de WooCommerce ("Actualizar productos existentes").
