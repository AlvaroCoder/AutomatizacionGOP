package teucost.modelos;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import teucost.modelos.excels.ExcelConciliado;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class LecturaConciliado {

    // =========================================================
    //  CONFIGURACIÓN
    // =========================================================

    private final ExcelConciliado excelConfig;
    private String nombreHoja = "Cont-1-26";

    // =========================================================
    //  CACHÉ — visita → DatosConciliado
    // =========================================================

    private Map<String, ExcelConciliado.DatosConciliado> cachePorVisita = null;

    // =========================================================
    //  CONSTRUCTOR
    // =========================================================

    public LecturaConciliado(String rutaExcel) {
        this.excelConfig = new ExcelConciliado(rutaExcel);
    }

    // =========================================================
    //  CARGA — una sola apertura del Excel
    // =========================================================

    /**
     * Abre el Excel, carga TODAS las filas en memoria y cierra.
     * Llamar UNA sola vez antes de hacer consultas.
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
            cargarFilas(hoja, null);
            System.out.println("[OK] Conciliado cargado — registros: " + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] LecturaConciliado.cargarTodosLosDatos: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    /**
     * Igual que cargarTodosLosDatos pero filtra solo las visitas indicadas.
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
            cargarFilas(hoja, filtro);
            System.out.println("[OK] Conciliado cargado para " + visitasFiltro.size()
                    + " visitas — encontradas: " + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] LecturaConciliado.cargarDatosDeVisitas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    // =========================================================
    //  CARGA DE FILAS EN CACHÉ
    // =========================================================

    private void cargarFilas(XSSFSheet hoja, Set<String> filtro) {
        int total = 0;
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            XSSFRow row = hoja.getRow(i);
            if (row == null) continue;

            String visita = getCellString(row.getCell(excelConfig.getColVisita())).trim();
            if (visita.isEmpty() || visita.equals("V.V.")) continue;
            if (filtro != null && !filtro.contains(visita)) continue;

            String muelle = getCellString(row.getCell(excelConfig.getColMuelle())).trim();

            cachePorVisita.put(visita,
                    new ExcelConciliado.DatosConciliado(visita, muelle));
            total++;
        }
        System.out.println("[INFO] Conciliado — filas cargadas: " + total);
    }

    // =========================================================
    //  CONSULTAS
    // =========================================================

    /**
     * Retorna el muelle de una visita, o cadena vacía si no se encuentra.
     */
    public String extraerMuelleNave(String nroVisita) {
        verificarCache();
        ExcelConciliado.DatosConciliado datos = cachePorVisita.get(nroVisita);
        if (datos == null) {
            System.out.println("[WARN] Conciliado — visita no encontrada: " + nroVisita);
            return "";
        }
        return datos.muelle;
    }

    /**
     * Retorna el registro completo de una visita.
     */
    public ExcelConciliado.DatosConciliado obtenerDatos(String nroVisita) {
        verificarCache();
        return cachePorVisita.getOrDefault(nroVisita, null);
    }

    // =========================================================
    //  UTILIDADES PRIVADAS
    // =========================================================

    private void verificarCache() {
        if (cachePorVisita == null) {
            throw new IllegalStateException(
                    "Caché vacío — llama a cargarTodosLosDatos() o cargarDatosDeVisitas() primero.");
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

    public void setNombreHoja(String nombreHoja) { this.nombreHoja = nombreHoja; }
}