package teucost.modelos;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import teucost.modelos.excels.ExcelDatosNave;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.time.temporal.WeekFields;
import java.util.*;

/**
 * Lee el archivo Control Cosmos UNA sola vez y responde consultas
 * para cualquier número de visitas sin volver a abrir el Excel.
 *
 * Flujo de uso:
 *   LecturaControlCosmos lcc = new LecturaControlCosmos(ruta);
 *   lcc.cargarTodosLosDatos();
 *   ExcelDatosNave.DatosNave datos = lcc.obtenerDatosNave("058-26");
 */
public class LecturaControlCosmos {

    // =========================================================
    //  CONFIGURACIÓN
    // =========================================================

    private final ExcelDatosNave excelConfig;
    private String nombreHoja = "Resumen Visitas Feb 26";

    private static final DateTimeFormatter FORMATTER =
            DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);



    private Map<String, ExcelDatosNave.DatosNave> cachePorVisita = null;

    public LecturaControlCosmos(String rutaExcel) {
        this.excelConfig = new ExcelDatosNave(rutaExcel);
    }

    /**
     * Abre el Excel, detecta columnas por encabezado, carga TODAS las
     * visitas en memoria y cierra el archivo.
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

            detectarColumnas(hoja);
            cargarFilas(hoja, null);

            System.out.println("[OK] ControlCosmos cargado — visitas encontradas: "
                    + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] cargarTodosLosDatos: " + e.getMessage());
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

            detectarColumnas(hoja);
            cargarFilas(hoja, filtro);

            System.out.println("[OK] ControlCosmos cargado para " + visitasFiltro.size()
                    + " visitas — encontradas: " + cachePorVisita.size());

        } catch (Exception e) {
            System.out.println("[ERROR] cargarDatosDeVisitas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    // =========================================================
    //  DETECCIÓN DE COLUMNAS — dinámica por encabezado
    // =========================================================

    private void detectarColumnas(XSSFSheet hoja) {
        XSSFRow encabezado = hoja.getRow(0);
        if (encabezado == null) {
            throw new RuntimeException("No se encontró fila de encabezado en la hoja.");
        }

        Map<String, Integer> encontradas = new HashMap<>();
        for (Cell cell : encabezado) {
            if (cell.getCellType() != CellType.STRING) continue;
            String valor = cell.getStringCellValue().trim();
            switch (valor) {
                case "Vessel Name": encontradas.put(excelConfig.getKeyColVesselName(), cell.getColumnIndex()); break;
                case "Line":        encontradas.put(excelConfig.getKeyColLine(),        cell.getColumnIndex()); break;
                case "Service":     encontradas.put(excelConfig.getKeyColService(),     cell.getColumnIndex()); break;
                case "Start Work":  encontradas.put(excelConfig.getKeyColStartWork(),   cell.getColumnIndex()); break;
                case "End Work":    encontradas.put(excelConfig.getKeyColEndWork(),      cell.getColumnIndex()); break;
                case "Cuadrillas":  encontradas.put(excelConfig.getKeyColCuadrillas(),  cell.getColumnIndex()); break;
                case "Visit":       encontradas.put(excelConfig.getKeyColVisit(),        cell.getColumnIndex()); break;
            }
        }

        validarColumna(encontradas, excelConfig.getKeyColVesselName());
        validarColumna(encontradas, excelConfig.getKeyColLine());
        validarColumna(encontradas, excelConfig.getKeyColService());
        validarColumna(encontradas, excelConfig.getKeyColStartWork());
        validarColumna(encontradas, excelConfig.getKeyColEndWork());
        validarColumna(encontradas, excelConfig.getKeyColCuadrillas());
        validarColumna(encontradas, excelConfig.getKeyColVisit());

        excelConfig.setColVesselName(  encontradas.get(excelConfig.getKeyColVesselName()));
        excelConfig.setColLine(        encontradas.get(excelConfig.getKeyColLine()));
        excelConfig.setColService(     encontradas.get(excelConfig.getKeyColService()));
        excelConfig.setColStartWork(   encontradas.get(excelConfig.getKeyColStartWork()));
        excelConfig.setColEndWork(     encontradas.get(excelConfig.getKeyColEndWork()));
        excelConfig.setColCuadrillas(  encontradas.get(excelConfig.getKeyColCuadrillas()));
        excelConfig.setColVisit(       encontradas.get(excelConfig.getKeyColVisit()));

        System.out.println("[OK] Columnas detectadas correctamente en ControlCosmos.");
    }

    private void validarColumna(Map<String, Integer> map, String key) {
        if (!map.containsKey(key)) {
            throw new RuntimeException("Columna requerida no encontrada: '" + key + "'");
        }
    }

    // =========================================================
    //  CARGA DE FILAS EN CACHÉ
    // =========================================================

    /**
     * Recorre todas las filas de datos y construye un DatosNave por visita.
     *
     * @param filtro Si es null carga todas las visitas.
     *               Si tiene valores, carga solo las visitas contenidas en él.
     */
    private void cargarFilas(XSSFSheet hoja, Set<String> filtro) {
        int totalCargadas = 0;
        // Fila 0 = encabezado, datos desde fila 1
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            XSSFRow row = hoja.getRow(i);
            if (row == null) continue;

            String visita = getCellString(row.getCell(excelConfig.getColVisit())).trim();
            if (visita.isEmpty()) continue;
            if (filtro != null && !filtro.contains(visita)) continue;

            String vesselName  = getCellString(row.getCell(excelConfig.getColVesselName()));
            String line        = getCellString(row.getCell(excelConfig.getColLine()));
            String service     = getCellString(row.getCell(excelConfig.getColService()));
            String startWork   = getCellString(row.getCell(excelConfig.getColStartWork()));
            String endWork     = getCellString(row.getCell(excelConfig.getColEndWork()));
            String cuadrillas  = getCellString(row.getCell(excelConfig.getColCuadrillas()));

            ExcelDatosNave.DatosNave datosNave = new ExcelDatosNave.DatosNave(
                    visita, vesselName, line, service, startWork, endWork, cuadrillas
            );

            cachePorVisita.put(visita, datosNave);
            totalCargadas++;
        }
        System.out.println("[INFO] Filas ControlCosmos cargadas en caché: " + totalCargadas);
    }

    // =========================================================
    //  CONSULTAS — todas reciben nroVisita como parámetro
    // =========================================================

    /**
     * Retorna el DatosNave crudo para una visita.
     * Útil cuando el llamador necesita todos los campos directamente.
     */
    public ExcelDatosNave.DatosNave obtenerDatosNave(String nroVisita) {
        verificarCache();
        ExcelDatosNave.DatosNave datos = cachePorVisita.get(nroVisita);
        if (datos == null) {
            System.out.println("[WARN] Visita no encontrada en ControlCosmos: " + nroVisita);
        }
        return datos;
    }

    /**
     * Retorna los datos calculados de una visita: semana ISO, tiempo efectivo
     * en horas, y todos los campos de DatosNave, empaquetados en un HashMap
     * con las mismas claves que usaba el sistema anterior para compatibilidad.
     */
    public HashMap<String, String> extraerResumenNaveCuad(String nroVisita) {
        verificarCache();
        HashMap<String, String> resumen = new HashMap<>();
        ExcelDatosNave.DatosNave datos = cachePorVisita.get(nroVisita);

        if (datos == null) {
            System.out.println("[WARN] extraerResumenNaveCuad — visita no encontrada: " + nroVisita);
            return resumen;
        }

        String nroSemana       = "";
        String tiempoEfectivo  = "";

        try {
            LocalDateTime dtInicio = LocalDateTime.parse(datos.startWork.trim(), FORMATTER);
            LocalDateTime dtFin    = LocalDateTime.parse(datos.endWork.trim(),   FORMATTER);

            int semana = dtFin.toLocalDate().get(WeekFields.ISO.weekOfWeekBasedYear());
            nroSemana  = String.valueOf(semana);

            long   minutos = ChronoUnit.MINUTES.between(dtInicio, dtFin);
            double horas   = minutos / 60.0;
            tiempoEfectivo = String.format("%.2f", horas);

        } catch (Exception e) {
            System.out.println("[WARN] No se pudo parsear fechas para visita "
                    + nroVisita + " → " + e.getMessage());
        }

        resumen.put("nombreVisita",    datos.vesselName);
        resumen.put("lineaServicio",   datos.lineVessel + "-" + datos.service);
        resumen.put("cuadrillas",      datos.cuadrillas);
        resumen.put("fechaFinalizacion", datos.endWork);
        resumen.put("fechaInicio",     datos.startWork);
        resumen.put("semana",          nroSemana);
        resumen.put("tiempoEfectivo",  tiempoEfectivo);

        return resumen;
    }

    /**
     * Retorna el número de mes (1-12) de la fecha de fin de una visita.
     * Retorna -1 si la visita no existe o la fecha no es parseable.
     */
    public int extraerNumeroMes(String nroVisita) {
        verificarCache();
        ExcelDatosNave.DatosNave datos = cachePorVisita.get(nroVisita);
        if (datos == null) return -1;
        try {
            return LocalDateTime.parse(datos.endWork.trim(), FORMATTER).getMonthValue();
        } catch (Exception e) {
            System.out.println("[WARN] No se pudo extraer mes para visita: " + nroVisita);
            return -1;
        }
    }

    /**
     * Retorna el mes abreviado en inglés en mayúsculas ("JAN", "FEB", etc.)
     * de la fecha de fin de una visita.
     */
    public String extraerMes(String nroVisita) {
        verificarCache();
        ExcelDatosNave.DatosNave datos = cachePorVisita.get(nroVisita);
        if (datos == null) return "";
        try {
            LocalDateTime dt = LocalDateTime.parse(datos.endWork.trim(), FORMATTER);
            return dt.getMonth()
                    .getDisplayName(java.time.format.TextStyle.SHORT, Locale.ENGLISH)
                    .toUpperCase();
        } catch (Exception e) {
            System.out.println("[WARN] No se pudo extraer mes para visita: " + nroVisita);
            return "";
        }
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