package teucost.modelos;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import teucost.modelos.excels.ExcelThroughput;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Lee el archivo Throughput UNA sola vez y responde consultas
 * de movimientos y TEUs por visita sin volver a abrir el Excel.
 *
 * NOTA SOBRE LA ESTRUCTURA DEL EXCEL:
 * El Throughput está transpuesto respecto a las otras fuentes:
 *   - La fila 8 (filaVisita) contiene los nros de visita como encabezados de columna
 *   - Cada columna = una visita
 *   - Las filas 30-37 (filaMovIni–filaMovFin) contienen movimientos de contenedores
 *   - La fila 39 (filaTeus) contiene el total de TEUs
 *
 * Flujo de uso:
 *   LecturaThroughput lt = new LecturaThroughput(ruta);
 *   lt.cargarDatosDeVisitas(listaVisitas);
 *   HashMap<String,Integer> res = lt.extraerResumenMovTeus("058-26");
 */
public class LecturaThroughput {

    // =========================================================
    //  CONFIGURACIÓN
    // =========================================================

    private final ExcelThroughput excelConfig;
    private String nombreHoja = "Flash Report";

    // =========================================================
    //  CACHÉ — visita → DatosThroughput
    // =========================================================

    private Map<String, ExcelThroughput.DatosThroughput> cachePorVisita = null;

    // =========================================================
    //  CONSTRUCTORES
    // =========================================================

    public LecturaThroughput() {
        this.excelConfig = new ExcelThroughput("");
    }

    public LecturaThroughput(String rutaThroughput) {
        this.excelConfig = new ExcelThroughput(rutaThroughput);
    }

    // =========================================================
    //  CARGA — una sola apertura del Excel
    // =========================================================

    /**
     * Carga solo las visitas indicadas (recomendado — el Throughput
     * puede tener muchas columnas con datos históricos).
     */
    public void cargarDatosDeVisitas(List<String> visitasFiltro) {
        cachePorVisita = new HashMap<>();
        Set<String> filtro = new HashSet<>(visitasFiltro);
        XSSFWorkbook wb = null;
        try {
            wb = abrirWorkbook(excelConfig.getRutaExcel());
            XSSFSheet hoja = wb.getSheet(this.nombreHoja);
            if (hoja == null) {
                throw new RuntimeException("Hoja no encontrada: '" + this.nombreHoja + "'");
            }
            cargarColumnas(hoja, filtro);
            System.out.println("[OK] Throughput cargado para " + visitasFiltro.size()
                    + " visitas — encontradas: " + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] LecturaThroughput.cargarDatosDeVisitas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    /**
     * Carga todas las visitas presentes en el archivo.
     */
    public void cargarTodosLosDatos() {
        cachePorVisita = new HashMap<>();
        XSSFWorkbook wb = null;
        try {
            wb = abrirWorkbook(excelConfig.getRutaExcel());
            XSSFSheet hoja = wb.getSheet(this.nombreHoja);
            if (hoja == null) {
                throw new RuntimeException("Hoja no encontrada: '" + this.nombreHoja + "'");
            }
            cargarColumnas(hoja, null);
            System.out.println("[OK] Throughput cargado — visitas: " + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] LecturaThroughput.cargarTodosLosDatos: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    // =========================================================
    //  CARGA DE COLUMNAS EN CACHÉ
    // =========================================================

    /**
     * Recorre la fila de encabezado (filaVisita) buscando nros de visita,
     * y para cada columna que coincida con el filtro extrae movimientos y TEUs.
     */
    private void cargarColumnas(XSSFSheet hoja, Set<String> filtro) {
        XSSFRow filaEncabezado = hoja.getRow(excelConfig.getFilaVisita());
        if (filaEncabezado == null) {
            System.out.println("[WARN] Throughput — fila de visitas vacía (fila "
                    + excelConfig.getFilaVisita() + ")");
            return;
        }

        int totalCargadas = 0;
        int lastCol = filaEncabezado.getLastCellNum();

        for (int col = 0; col < lastCol; col++) {
            String visita = getCellString(filaEncabezado.getCell(col)).trim();
            if (visita.isEmpty()) continue;
            if (filtro != null && !filtro.contains(visita)) continue;

            // ── Movimientos de contenedores (suma de filaMovIni a filaMovFin) ──
            int sumaMov = 0;
            for (int fila = excelConfig.getFilaMovIni();
                 fila <= excelConfig.getFilaMovFin(); fila++) {
                XSSFRow row = hoja.getRow(fila);
                if (row == null) continue;
                String val = getCellString(row.getCell(col)).trim();
                if (!val.isEmpty()) {
                    try {
                        sumaMov += (int) Double.parseDouble(val.replace(",", "."));
                    } catch (NumberFormatException ignored) {}
                }
            }

            // ── TEUs (fila filaTeus) ──────────────────────────────
            int teus = 0;
            XSSFRow filaTeus = hoja.getRow(excelConfig.getFilaTeus());
            if (filaTeus != null) {
                String teusStr = getCellString(filaTeus.getCell(col)).trim();
                if (!teusStr.isEmpty()) {
                    try {
                        teus = (int) Double.parseDouble(teusStr.replace(",", "."));
                    } catch (NumberFormatException ignored) {}
                }
            }

            cachePorVisita.put(visita,
                    new ExcelThroughput.DatosThroughput(visita, sumaMov, teus));
            totalCargadas++;
        }

        System.out.println("[INFO] Throughput — columnas cargadas: " + totalCargadas);
    }

    // =========================================================
    //  CONSULTAS
    // =========================================================

    /**
     * Retorna un mapa con "sumaMovContenedores" y "sumaTeus" para la visita.
     * Compatible con el contrato que espera ControladorCostoTeu.
     */
    public HashMap<String, Integer> extraerResumenMovTeus(String nroVisita) {
        verificarCache();
        HashMap<String, Integer> resumen = new HashMap<>();
        resumen.put("sumaMovContenedores", 0);
        resumen.put("sumaTeus",            0);

        ExcelThroughput.DatosThroughput datos = cachePorVisita.get(nroVisita);
        if (datos == null) {
            System.out.println("[WARN] Throughput — visita no encontrada: " + nroVisita);
            return resumen;
        }

        resumen.put("sumaMovContenedores", datos.movContenedores);
        resumen.put("sumaTeus",            datos.teus);
        return resumen;
    }

    /**
     * Retorna el registro completo de una visita.
     */
    public ExcelThroughput.DatosThroughput obtenerDatos(String nroVisita) {
        verificarCache();
        return cachePorVisita.getOrDefault(nroVisita, null);
    }

    // =========================================================
    //  UTILIDADES PRIVADAS
    // =========================================================

    private void verificarCache() {
        if (cachePorVisita == null) {
            throw new IllegalStateException(
                    "Caché vacío — llama a cargarDatosDeVisitas() o cargarTodosLosDatos() primero.");
        }
    }

    private XSSFWorkbook abrirWorkbook(String ruta) throws IOException {
        try (InputStream fis = new FileInputStream(new File(ruta))) {
            return new XSSFWorkbook(fis);
        }
    }

    private void cerrarWorkbook(XSSFWorkbook wb) {
        if (wb != null) try { wb.close(); } catch (Exception ignored) {}
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                return d == Math.floor(d) ? String.valueOf((long) d) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:  return cell.getRichStringCellValue().getString().trim();
                    case NUMERIC:
                        double df = cell.getNumericCellValue();
                        return df == Math.floor(df) ? String.valueOf((long) df) : String.valueOf(df);
                    default: return "";
                }
            default: return "";
        }
    }

    // =========================================================
    //  SETTERS
    // =========================================================

    public void setNombreHoja(String nombreHoja)   { this.nombreHoja = nombreHoja; }
    public void setFilaVisita(int filaVisita)       { excelConfig.setFilaVisita(filaVisita); }
    public void setFilaMovIni(int filaMovIni)       { excelConfig.setFilaMovIni(filaMovIni); }
    public void setFilaMovFin(int filaMovFin)       { excelConfig.setFilaMovFin(filaMovFin); }
    public void setFilaTeus(int filaTeus)           { excelConfig.setFilaTeus(filaTeus);     }
}