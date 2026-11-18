import { stat } from "node:fs/promises";
import path from "node:path";
import XLSX from "xlsx";

const EXTENSIONES = new Set([".xlsx", ".xlsm", ".xls"]);

export type Celda = string | number | boolean | Date | null;
export type Fila = Record<string, Celda>;

export interface Tabla {
  hoja: string;
  columnas: string[];
  filas: Fila[];
}

export async function rutaExcel(archivo: string): Promise<string> {
  const ruta = path.resolve(archivo);
  if (!EXTENSIONES.has(path.extname(ruta).toLowerCase())) {
    throw new Error("El archivo debe tener extension .xlsx, .xlsm o .xls");
  }
  const info = await stat(ruta);
  if (!info.isFile()) throw new Error("La ruta no apunta a un archivo");
  return ruta;
}

export function leerLibro(ruta: string): XLSX.WorkBook {
  const libro = XLSX.readFile(ruta, { cellDates: true });
  if (libro.SheetNames.length === 0) throw new Error("El libro no tiene hojas");
  return libro;
}

export function seleccionarHoja(libro: XLSX.WorkBook, nombre?: string): string {
  if (!nombre) return libro.SheetNames[0];
  const encontrada = libro.SheetNames.find((h) => h.toLowerCase() === nombre.toLowerCase());
  if (!encontrada) {
    throw new Error(`No existe la hoja '${nombre}'. Hojas disponibles: ${libro.SheetNames.join(", ")}`);
  }
  return encontrada;
}

export function esVacio(valor: unknown): boolean {
  return valor === null || valor === undefined || (typeof valor === "string" && valor.trim() === "");
}

function normalizar(valor: unknown): Celda {
  if (esVacio(valor)) return null;
  if (typeof valor === "number" || typeof valor === "string" || typeof valor === "boolean" || valor instanceof Date) {
    return valor;
  }
  return String(valor);
}

export function obtenerTabla(libro: XLSX.WorkBook, hoja: string): Tabla {
  const matriz = XLSX.utils.sheet_to_json<unknown[]>(libro.Sheets[hoja], {
    header: 1,
    defval: null,
    blankrows: false,
    raw: true,
  });
  const vistos = new Map<string, number>();
  const columnas = ((matriz[0] ?? []) as unknown[]).map((valor, indice) => {
    const base = esVacio(valor) ? `Columna_${indice + 1}` : String(valor).trim();
    const repeticiones = (vistos.get(base) ?? 0) + 1;
    vistos.set(base, repeticiones);
    return repeticiones === 1 ? base : `${base}_${repeticiones}`;
  });
  if (columnas.length === 0) throw new Error(`La hoja '${hoja}' no tiene encabezados`);

  const filas = matriz.slice(1).map((celdas) => {
    const valores = (celdas ?? []) as unknown[];
    const fila: Fila = {};
    columnas.forEach((columna, indice) => {
      fila[columna] = normalizar(valores[indice]);
    });
    return fila;
  });
  return { hoja, columnas, filas };
}

export function numero(valor: unknown): number | null {
  if (esVacio(valor)) return null;
  if (typeof valor === "number") return Number.isFinite(valor) ? valor : null;
  if (typeof valor === "string") {
    const numero = Number(valor.trim());
    return Number.isFinite(numero) ? numero : null;
  }
  return null;
}

export function redondear(valor: number): number {
  return Number(valor.toFixed(4));
}

export function escribirHoja(
  libro: XLSX.WorkBook,
  nombre: string,
  columnas: string[],
  filas: Fila[],
): void {
  if (libro.SheetNames.includes(nombre)) {
    delete libro.Sheets[nombre];
    libro.SheetNames = libro.SheetNames.filter((hoja) => hoja !== nombre);
  }
  const matriz = [columnas, ...filas.map((fila) => columnas.map((columna) => fila[columna] ?? null))];
  XLSX.utils.book_append_sheet(libro, XLSX.utils.aoa_to_sheet(matriz), nombre, true);
}
