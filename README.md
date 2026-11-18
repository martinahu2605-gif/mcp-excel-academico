# MCP Excel Académico

Servidor **MCP** (Model Context Protocol) en TypeScript que da a un asistente de IA tres herramientas para trabajar con planillas de cálculo locales: inspeccionarlas, limpiar filas incompletas y generar un resumen estadístico. Está pensado para docencia e investigación académica y funciona con opencode, Claude Desktop, Cursor o cualquier cliente MCP por stdio.

## 1. Descripción y problema académico

En trabajos de laboratorio, estadística o investigación clínica es habitual recibir un Excel con:

- celdas vacías que distorsionan promedios y conteos;
- hojas con decenas de columnas cuyo significado no está documentado;
- la necesidad de repetir a mano la misma limpieza y el mismo resumen para cada entrega.

Hacerlo manualmente es lento y difícil de reproducir: dos personas pueden obtener resultados distintos aplicando criterios diferentes. Este MCP convierte esas tres tareas en herramientas explícitas, con parámetros validados y resultados verificables, para que el asistente pueda ayudar sin editar el archivo "a ciegas". Los datos se procesan **localmente**, sin subirlos a ningún servicio externo.

El flujo recomendado es siempre: **inspeccionar → limpiar → resumir**.

## 2. Requisitos

- **Node.js 18 o superior** y `npm` (ver `node --version`).
- Un **cliente MCP**: [opencode](https://opencode.ai) (recomendado), Claude Desktop, Cursor o el MCP Inspector.
- Un archivo de datos local (en este repo, `datos_ejemplo.xlsx`).
- No hace falta tener Microsoft Excel instalado: la lectura y escritura se hacen con la librería `xlsx` (SheetJS).

Funciona en Windows, macOS y Linux. En Windows, si PowerShell bloquea `npm.ps1` por su política de scripts, usa `npm.cmd` en todos los comandos.

## 3. Instalación paso a paso

```bash
git clone https://github.com/Maurichilean3d/mcp-starter-kit.git mcp-excel-academico
cd mcp-excel-academico
npm install
npm run build
npm test
```

`npm test` compila el código y ejecuta la prueba de humo; debe terminar con `Prueba de humo superada.` y listar las siete herramientas (las cuatro de la plantilla más las tres de Excel).

Después:

1. Abre tu cliente MCP **en esa misma carpeta** (`opencode`).
2. El archivo `opencode.json` ya registra el servidor; reinicia el cliente si ya estaba abierto antes de compilar.
3. Para probar sin conexión, ejecuta `npm run inspect` y llama a las herramientas desde el Inspector.

Si quieres volver a generar el archivo de ejemplo original (esto **sobrescribe** `datos_ejemplo.xlsx`):

```bash
node scripts/crear-datos-ejemplo.mjs
```

## 4. Archivo de ejemplo

`datos_ejemplo.xlsx` contiene datos clínicos **ficticios** de laboratorio:

| Hoja | Contenido |
|---|---|
| `Laboratorio` | 20 registros, 13 columnas (identificador, fecha, género, edad, hemoglobina, hematocrito, leucocitos, plaquetas, glucosa, creatinina, sodio, potasio y diagnóstico). Tiene 7 celdas vacías repartidas en género, edad, plaquetas, glucosa, sodio y potasio. |
| `Referencias` | Rangos de referencia de los análisis. |

Al ejecutar las herramientas se añaden `Datos_Limpios` (13 filas sin celdas vacías) y `Resumen_Estadistico` (9 columnas numéricas).

## 5. Herramientas propias

| Herramienta | Qué hace | Escritura |
|---|---|---|
| `excel_inspeccionar` | Lee y describe la estructura de una hoja. | No (solo lectura) |
| `excel_limpiar_datos` | Elimina filas con celdas vacías, con filtro opcional, y guarda el resultado en `Datos_Limpios`. | Sí: exige `confirmar: true` |
| `excel_resumen_estadistico` | Calcula estadística descriptiva y la guarda en `Resumen_Estadistico`. | Sí: exige `confirmar: true` |

Las rutas se resuelven y validan antes de abrir el archivo; solo se aceptan extensiones `.xlsx`, `.xlsm` y `.xls`.

### 5.1. `excel_inspeccionar`

Devuelve: hojas del libro, hoja inspeccionada, columnas, total de filas, primeras 5 filas y cantidad de celdas vacías **por columna**. Úsala siempre antes de las otras dos.

**Parámetros**

- `archivo` (obligatorio): ruta al archivo Excel.
- `hoja` (opcional): nombre de la hoja; por defecto, la primera (`Laboratorio`).

**Ejemplo en lenguaje natural**

> *"Inspecciona datos_ejemplo.xlsx y dime qué columnas tienen datos incompletos."*

**Ejemplo de llamada**

```json
{
  "archivo": "C:\\mcp-excel-academico\\datos_ejemplo.xlsx"
}
```

**Resultado esperado (extracto)**

```json
{
  "hojas": ["Laboratorio", "Referencias"],
  "columnas": ["id_paciente", "fecha_muestra", "genero", "edad", "..."],
  "total_filas": 20,
  "celdas_vacias_por_columna": { "genero": 1, "edad": 1, "glucosa_mg_dL": 2 }
}
```

### 5.2. `excel_limpiar_datos`

Elimina **todas** las filas que tengan al menos una celda vacía. Opcionalmente filtra por una columna y un valor (comparación exacta, sin distinguir mayúsculas). Guarda el resultado en la hoja `Datos_Limpios` del **mismo archivo**, conservando la hoja original. Si la hoja destino ya existe, se sobrescribe.

**Parámetros**

- `archivo` (obligatorio): ruta al archivo Excel.
- `hoja` (opcional): hoja de origen; por defecto, la primera.
- `columna` y `valor` (opcionales, **deben ir juntos**): filtro a aplicar.
- `confirmar` (obligatorio): literal `true` para autorizar la escritura.

**Ejemplo en lenguaje natural**

> *"Limpia los datos incompletos de datos_ejemplo.xlsx y guárdalos en una hoja nueva; confirma la escritura."*

**Ejemplo de llamada**

```json
{
  "archivo": "C:\\mcp-excel-academico\\datos_ejemplo.xlsx",
  "confirmar": true
}
```

Para quedarse solo con las filas completas de un diagnóstico:

```json
{
  "archivo": "C:\\mcp-excel-academico\\datos_ejemplo.xlsx",
  "columna": "diagnostico",
  "valor": "Diabetes tipo 2",
  "confirmar": true
}
```

**Resultado esperado (extracto)**

```json
{
  "hoja_destino": "Datos_Limpios",
  "filas_originales": 20,
  "filas_eliminadas": 7,
  "filas_guardadas": 13,
  "filtro": null
}
```

### 5.3. `excel_resumen_estadistico`

Para cada columna cuyos valores no vacíos sean todos numéricos calcula: `conteo`, `promedio`, `desviacion_estandar`, `minimo` y `maximo`. La desviación es **muestral** (denominador `n - 1`) y los resultados se redondean a 4 decimales. Escribe una hoja `Resumen_Estadistico` con columnas `Columna`, `Conteo`, `Promedio`, `Desviacion_Estandar`, `Minimo` y `Maximo`.

**Parámetros**

- `archivo` (obligatorio): ruta al archivo Excel.
- `hoja` (opcional): hoja a analizar; por defecto, la primera. Para analizar los datos ya limpios, indica `"Datos_Limpios"`.
- `confirmar` (obligatorio): literal `true` para autorizar la escritura.

**Ejemplo en lenguaje natural**

> *"Genera el resumen estadístico de la hoja Datos_Limpios de datos_ejemplo.xlsx y confírmalo."*

**Ejemplo de llamada**

```json
{
  "archivo": "C:\\mcp-excel-academico\\datos_ejemplo.xlsx",
  "hoja": "Datos_Limpios",
  "confirmar": true
}
```

**Resultado esperado (extracto)**

```json
{
  "columnas_numericas": 9,
  "estadisticas": [
    {
      "columna": "hemoglobina_g_dL",
      "conteo": 13,
      "promedio": 13.05,
      "desviacion_estandar": 1.27,
      "minimo": 10.8,
      "maximo": 15
    }
  ]
}
```

> Los valores del extracto son ilustrativos; ejecuta la herramienta para obtener los cálculos reales sobre la hoja que elijas.

## 6. Responsabilidad, Seguridad y Límites conocidos

**Permisos y privacidad**

- El MCP solo accede a **archivos locales** cuya ruta proporciona el usuario; no abre conexiones de red ni envía datos a servicios externos.
- La validación de extensión admite actualmente `.xlsx`, `.xlsm` y `.xls`. **`.csv` forma parte del alcance previsto de la política de permisos, pero todavía no está implementado**: intentar usarlo devuelve un error controlado. No documents `.csv` como soportado hasta añadirlo al validador de `src/lib/excel.ts`.
- No se usan claves de API ni secretos en el código.
- Los datos de `datos_ejemplo.xlsx` son ficticios. Para datos reales de personas, anonimízalos y verifica que tienes autorización para procesarlos; este MCP no sustituye una revisión ética ni normativa.

**Confirmación y writing**

- `excel_limpiar_datos` y `excel_resumen_estadistico` declaran `destructiveHint: true` y exigen `confirmar: true`. Sin ese literal, devuelven un error y no tocan el archivo.
- Ambas escriben dentro del mismo libro y **sobrescriben** `Datos_Limpios` / `Resumen_Estadistico` si ya existen. No hay papelera, historial ni `undo`: haz una copia antes de ejecutarlas sobre un archivo importante.
- La limpieza **no elimina la hoja original**; solo añade una hoja nueva. Aun así, revisa el resultado antes de usarlo como fuente oficial.

**Límites conocidos**

- La primera fila se trata como encabezados; no se admiten tablas con encabezados en otra fila ni celdas combinadas complejas.
- Las filas totalmente vacías se ignoran; una fila se considera incompleta si **cualquier** columna está vacía (no hay umbral configurable).
- El filtro es de igualdad exacta (ignorando mayúsculas); no admite comodines, rangos ni expresiones regulares.
- No se evalúan fórmulas: se leen los valores guardados en caché por Excel.
- Una columna solo entra al resumen si todos sus valores no vacíos son numéricos o cadenas convertibles a número; las fechas y columnas mixtas se omiten.
- La desviación estándar es muestral; para población completa hay que ajustar el denominador.
- `hoja` selecciona una sola hoja por llamada; no se puede resumir el libro completo de una vez.
- Las extensiones `.xlsm` y `.xls` pasan el validador pero solo `.xlsx` se usa en las pruebas; las macros no se ejecutan ni se conservan de forma garantizada.
- No hay límite de tamaño de archivo ni bloqueos concurrentes: evita abrir el mismo libro en Excel mientras se escribe desde el MCP.

**Uso de IA generativa**

Este proyecto fue desarrollado con asistencia de **IA generativa** (opencode junto a un modelo de lenguaje) para redactar el código, las descripciones de las herramientas y esta documentación. El resultado fue compilado (`npm run build`), probado (`npm test`) y verificado manualmente ejecutando las tres herramientas sobre `datos_ejemplo.xlsx`. La IA es una herramienta de apoyo: las decisiones sobre qué datos limpiar, qué archivos autorizar y cómo interpretar los resultados son responsabilidad de quien usa el MCP.

## 7. Estructura relevante

```
src/index.ts                      Arranque; carga automática de src/tools/
src/tools/excel.ts                Las 3 herramientas de Excel
src/lib/excel.ts                  Validación de rutas, lectura/escritura y helpers
scripts/crear-datos-ejemplo.mjs   Generador del archivo de ejemplo
scripts/smoke-test.mjs            Prueba de humo del servidor
datos_ejemplo.xlsx                Libro de práctica
opencode.json                     Registro del MCP en opencode
docs/, AGENTS.md                  Guías y reglas de la plantilla original
```

## 8. Comandos útiles

| Comando | Qué hace |
|---|---|
| `npm install` | Instala dependencias (incluye `xlsx`). |
| `npm run build` | Compila TypeScript a `dist/`; obligatorio tras cambiar código. |
| `npm test` | Compila y verifica handshake, listado y ejecución de herramientas. |
| `npm run inspect` | Abre el MCP Inspector para probar las herramientas con interfaz visual. |
| `node scripts/crear-datos-ejemplo.mjs` | Recrea el Excel de ejemplo desde cero. |

## Licencia

MIT. Proyecto derivado de la plantilla `mcp-starter-kit`; conserva las herramientas de ejemplo `hora_actual`, `saludar`, `listar_aplicaciones` y `abrir_aplicacion`.
