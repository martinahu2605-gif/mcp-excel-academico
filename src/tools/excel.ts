import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import XLSX from "xlsx";
import { z } from "zod";
import { error, json } from "../lib/respuestas.js";
import {
  esVacio,
  escribirHoja,
  leerLibro,
  numero,
  obtenerTabla,
  redondear,
  rutaExcel,
  seleccionarHoja,
  type Celda,
  type Fila,
} from "../lib/excel.js";

function coincide(valorCelda: Celda, filtro: string | number): boolean {
  const numeroCelda = numero(valorCelda);
  const numeroFiltro = numero(filtro);
  if (numeroCelda !== null && numeroFiltro !== null) return numeroCelda === numeroFiltro;
  return String(valorCelda ?? "").trim().toLowerCase() === String(filtro).trim().toLowerCase();
}

function mensaje(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

export function register(server: McpServer): void {
  server.registerTool(
    "excel_inspeccionar",
    {
      title: "Inspeccionar archivo Excel",
      description:
        "Lee un archivo Excel (.xlsx/.xlsm/.xls) y devuelve sus hojas, columnas, total de filas, las primeras 5 filas y cuántas celdas vacías hay por columna. Usalo antes de limpiar o analizar datos para conocer su estructura.",
      inputSchema: {
        archivo: z.string().min(1).describe("Ruta al archivo Excel que se quiere inspeccionar"),
        hoja: z.string().optional().describe("Nombre de la hoja a inspeccionar; por defecto la primera"),
      },
      annotations: { readOnlyHint: true },
    },
    async ({ archivo, hoja }) => {
      try {
        const ruta = await rutaExcel(archivo);
        const libro = leerLibro(ruta);
        const nombreHoja = seleccionarHoja(libro, hoja);
        const tabla = obtenerTabla(libro, nombreHoja);
        const celdasVaciasPorColumna: Record<string, number> = {};
        for (const columna of tabla.columnas) {
          celdasVaciasPorColumna[columna] = tabla.filas.filter((fila) => esVacio(fila[columna])).length;
        }
        return json({
          archivo: ruta,
          hojas: libro.SheetNames,
          hoja_inspeccionada: nombreHoja,
          columnas: tabla.columnas,
          total_filas: tabla.filas.length,
          primeras_5_filas: tabla.filas.slice(0, 5),
          celdas_vacias_por_columna: celdasVaciasPorColumna,
        });
      } catch (e) {
        return error(`Fallo en excel_inspeccionar: ${mensaje(e)}`);
      }
    },
  );

  server.registerTool(
    "excel_limpiar_datos",
    {
      title: "Limpiar datos de Excel",
      description:
        "Elimina del archivo Excel las filas que tienen alguna celda vacía y, opcionalmente, filtra por una columna y un valor. Guarda el resultado en una hoja nueva llamada Datos_Limpios dentro del mismo archivo. Requiere confirmar: true porque sobrescribe esa hoja si ya existe.",
      inputSchema: {
        archivo: z.string().min(1).describe("Ruta al archivo Excel que se quiere limpiar"),
        hoja: z.string().optional().describe("Nombre de la hoja de origen; por defecto la primera"),
        columna: z.string().optional().describe("Columna por la que se quiere filtrar (se usa junto a valor)"),
        valor: z
          .union([z.string(), z.number()])
          .optional()
          .describe("Valor exacto a buscar en la columna; la comparación ignora mayúsculas"),
        confirmar: z.literal(true).describe("Debe ser true para autorizar la escritura de la hoja Datos_Limpios"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ archivo, hoja, columna, valor, confirmar }) => {
      if (!confirmar) return error("Falta confirmar: true para escribir en el archivo");
      if ((columna === undefined) !== (valor === undefined)) {
        return error("Las opciones columna y valor deben indicarse juntas o no indicarse ninguna");
      }
      try {
        const ruta = await rutaExcel(archivo);
        const libro = leerLibro(ruta);
        const nombreHoja = seleccionarHoja(libro, hoja);
        const tabla = obtenerTabla(libro, nombreHoja);
        let columnaFiltro: string | undefined;
        if (columna !== undefined) {
          columnaFiltro = tabla.columnas.find((c) => c.toLowerCase() === columna.toLowerCase());
          if (!columnaFiltro) {
            return error(`No existe la columna '${columna}'. Columnas: ${tabla.columnas.join(", ")}`);
          }
        }
        const filasLimpas = tabla.filas.filter((fila) => {
          if (tabla.columnas.some((c) => esVacio(fila[c]))) return false;
          if (columnaFiltro === undefined || valor === undefined) return true;
          return coincide(fila[columnaFiltro], valor);
        });
        escribirHoja(libro, "Datos_Limpios", tabla.columnas, filasLimpas);
        XLSX.writeFile(libro, ruta);
        return json({
          archivo: ruta,
          hoja_origen: nombreHoja,
          hoja_destino: "Datos_Limpios",
          filas_originales: tabla.filas.length,
          filas_eliminadas: tabla.filas.length - filasLimpas.length,
          filas_guardadas: filasLimpas.length,
          filtro: columnaFiltro ? { columna: columnaFiltro, valor } : null,
        });
      } catch (e) {
        return error(`Fallo en excel_limpiar_datos: ${mensaje(e)}`);
      }
    },
  );

  server.registerTool(
    "excel_resumen_estadistico",
    {
      title: "Resumen estadístico de Excel",
      description:
        "Calcula conteo, promedio, desviación estándar muestral, mínimo y máximo de todas las columnas numéricas de una hoja de Excel y crea la hoja Resumen_Estadistico en el mismo archivo. Requiere confirmar: true porque sobrescribe esa hoja si ya existe.",
      inputSchema: {
        archivo: z.string().min(1).describe("Ruta al archivo Excel que se quiere resumir"),
        hoja: z.string().optional().describe("Nombre de la hoja a analizar; por defecto la primera"),
        confirmar: z
          .literal(true)
          .describe("Debe ser true para autorizar la escritura de la hoja Resumen_Estadistico"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ archivo, hoja, confirmar }) => {
      if (!confirmar) return error("Falta confirmar: true para escribir en el archivo");
      try {
        const ruta = await rutaExcel(archivo);
        const libro = leerLibro(ruta);
        const nombreHoja = seleccionarHoja(libro, hoja);
        const tabla = obtenerTabla(libro, nombreHoja);
        const columnasResumen = [
          "Columna",
          "Conteo",
          "Promedio",
          "Desviacion_Estandar",
          "Minimo",
          "Maximo",
        ];
        const filasResumen: Fila[] = [];
        const estadisticas: Record<string, unknown>[] = [];
        for (const columna of tabla.columnas) {
          const valores = tabla.filas
            .map((fila) => numero(fila[columna]))
            .filter((valor): valor is number => valor !== null);
          const noVacios = tabla.filas.filter((fila) => !esVacio(fila[columna])).length;
          if (valores.length === 0 || valores.length !== noVacios) continue;
          const promedio = valores.reduce((suma, valor) => suma + valor, 0) / valores.length;
          const varianza =
            valores.length > 1
              ? valores.reduce((suma, valor) => suma + (valor - promedio) ** 2, 0) / (valores.length - 1)
              : 0;
          const resumen = {
            columna,
            conteo: valores.length,
            promedio: redondear(promedio),
            desviacion_estandar: redondear(Math.sqrt(varianza)),
            minimo: redondear(Math.min(...valores)),
            maximo: redondear(Math.max(...valores)),
          };
          estadisticas.push(resumen);
          filasResumen.push({
            Columna: resumen.columna,
            Conteo: resumen.conteo,
            Promedio: resumen.promedio,
            Desviacion_Estandar: resumen.desviacion_estandar,
            Minimo: resumen.minimo,
            Maximo: resumen.maximo,
          });
        }
        escribirHoja(libro, "Resumen_Estadistico", columnasResumen, filasResumen);
        XLSX.writeFile(libro, ruta);
        return json({
          archivo: ruta,
          hoja_origen: nombreHoja,
          hoja_destino: "Resumen_Estadistico",
          columnas_numericas: estadisticas.length,
          estadisticas,
        });
      } catch (e) {
        return error(`Fallo en excel_resumen_estadistico: ${mensaje(e)}`);
      }
    },
  );
}
