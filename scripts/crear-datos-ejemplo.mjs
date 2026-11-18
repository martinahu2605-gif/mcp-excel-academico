import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

const ruta = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "datos_ejemplo.xlsx");

const encabezados = [
  "id_paciente",
  "fecha_muestra",
  "genero",
  "edad",
  "hemoglobina_g_dL",
  "hematocrito_pct",
  "leucocitos_mil_uL",
  "plaquetas_mil_uL",
  "glucosa_mg_dL",
  "creatinina_mg_dL",
  "sodio_mEq_L",
  "potasio_mEq_L",
  "diagnostico",
];

const registros = [
  ["P-001", "2026-10-02", "F", 34, 13.2, 39.8, 6.7, 245, 88, 0.78, 139, 4.1, "Control rutina"],
  ["P-002", "2026-10-02", "M", 58, 11.4, 34.1, 8.2, 198, 112, 1.12, 141, 4.4, "Anemia ferropenica"],
  ["P-003", "2026-10-03", "F", 47, 13.8, 41.2, 5.9, 268, "", 0.84, 137, 3.9, "Diabetes tipo 2"],
  ["P-004", "2026-10-03", "M", 63, 12.6, 37.5, 7.4, 231, 104, 0.96, 140, 4.6, "Hipertension arterial"],
  ["P-005", "2026-10-04", "F", 29, 14.1, 42.3, 5.2, 302, 82, 0.71, 138, 4.0, "Control prenatal"],
  ["P-006", "2026-10-04", "M", 71, 10.8, 32.6, 9.1, 176, 134, 1.48, 136, "", "Insuficiencia renal leve"],
  ["P-007", "2026-10-05", "F", 45, 12.9, 38.7, 6.3, "", 95, 0.83, 142, 4.2, "Hipotiroidismo"],
  ["P-008", "2026-10-05", "M", 39, 15.0, 44.9, 5.7, 289, 91, 0.76, 139, 4.3, "Control rutina"],
  ["P-009", "2026-10-06", "F", 52, 11.9, 35.4, 10.3, 212, 148, 0.92, 141, 4.7, "Infeccion respiratoria"],
  ["P-010", "2026-10-06", "M", 66, 13.5, 40.1, 7.8, 254, 106, 1.05, "", 4.5, "Diabetes tipo 2"],
  ["P-011", "2026-10-07", "F", "", 12.4, 36.9, 6.1, 223, 99, 0.88, 137, 3.8, "Anemia ferropenica"],
  ["P-012", "2026-10-07", "M", 41, 14.6, 43.5, 5.4, 276, 86, 0.73, 140, 4.2, "Control rutina"],
  ["P-013", "2026-10-08", "F", 37, 13.1, 39.2, 8.7, 241, 93, 0.81, 138, 4.0, "Infeccion urinaria"],
  ["P-014", "2026-10-08", "", 59, 11.7, 34.8, 7.1, 187, 118, 1.21, 142, 4.6, "Hipertension arterial"],
  ["P-015", "2026-10-09", "F", 48, 12.8, 38.1, 6.5, 259, "", 0.90, 139, 4.1, "Control rutina"],
  ["P-016", "2026-10-09", "M", 54, 10.9, 33.2, 11.2, 165, 127, 1.34, 135, 5.1, "Insuficiencia renal leve"],
  ["P-017", "2026-10-10", "F", 32, 14.4, 43.1, 5.6, 291, 79, 0.69, 141, 3.9, "Control prenatal"],
  ["P-018", "2026-10-10", "M", 68, 12.2, 36.4, 8.4, 205, 109, 1.17, 138, 4.4, "Diabetes tipo 2"],
  ["P-019", "2026-10-11", "F", 61, 13.7, 40.8, 6.9, 248, 97, 0.87, 140, 4.2, "Hipertension arterial"],
  ["P-020", "2026-10-11", "M", 46, 14.9, 44.2, 5.8, 273, 84, 0.75, 137, 4.0, "Control rutina"],
];

const referencias = [
  ["Analito", "Rango de referencia", "Unidad"],
  ["Hemoglobina", "12.0 - 16.0", "g/dL"],
  ["Hematocrito", "36.0 - 46.0", "%"],
  ["Leucocitos", "4.0 - 11.0", "x10^3/uL"],
  ["Plaquetas", "150 - 400", "x10^3/uL"],
  ["Glucosa", "70 - 99", "mg/dL"],
  ["Creatinina", "0.6 - 1.1", "mg/dL"],
  ["Sodio", "135 - 145", "mEq/L"],
  ["Potasio", "3.5 - 5.0", "mEq/L"],
];

const libro = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(libro, XLSX.utils.aoa_to_sheet([encabezados, ...registros]), "Laboratorio");
XLSX.utils.book_append_sheet(libro, XLSX.utils.aoa_to_sheet(referencias), "Referencias");
XLSX.writeFile(libro, ruta);
console.log(`Creado ${ruta} con ${registros.length} registros y ${referencias.length - 1} rangos de referencia.`);
