package teucost.modelos;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import teucost.modelos.excels.ExcelMoveHistoryTeuCost;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.*;


public class LecturaMoveHistory {
    private final ExcelMoveHistoryTeuCost excelConfig;
    private String nombreHoja = "MoveEvent";

    private static final DateTimeFormatter FORMATTER =
            DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);

    private static final int UMBRAL_BREAK_MINUTOS = 15;
    private Map<String, Map<String, List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost>>>
            cachePorVisitaYGrua = null;
    public LecturaMoveHistory(String rutaExcel) {
        this.excelConfig = new ExcelMoveHistoryTeuCost(rutaExcel);
    }
    public void cargarTodosLosDatos() {
        cachePorVisitaYGrua = new HashMap<>();
        XSSFWorkbook wb = null;
        try {
            wb = abrirWorkbook(excelConfig.getRutaExcel());
            XSSFSheet hoja = wb.getSheet(this.nombreHoja);
            if (hoja == null) {
                throw new RuntimeException("Hoja no encontrada: '" + this.nombreHoja + "'");
            }

            detectarColumnas(hoja);
            cargarFilas(hoja, null);

            System.out.println("[OK] MoveHistory cargado — visitas encontradas: "
                    + cachePorVisitaYGrua.size());

        } catch (Exception e) {
            System.out.println("[ERROR] cargarTodosLosDatos: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

    public void cargarDatosDeVisitas(List<String> visitasFiltro) {
        cachePorVisitaYGrua = new HashMap<>();
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

            System.out.println("[OK] MoveHistory cargado para " + visitasFiltro.size()
                    + " visitas — encontradas: " + cachePorVisitaYGrua.size());

        } catch (Exception e) {
            System.out.println("[ERROR] cargarDatosDeVisitas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            cerrarWorkbook(wb);
        }
    }

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
                case "Time Completed":  encontradas.put(excelConfig.getKeyColTimeCompleted(),  cell.getColumnIndex()); break;
                case "Crane CHE Name":  encontradas.put(excelConfig.getKeyColCraneCheName(),   cell.getColumnIndex()); break;
                case "Fetch CHE Name":  encontradas.put(excelConfig.getKeyColFetchCheName(),   cell.getColumnIndex()); break;
                case "Carry CHE Name":  encontradas.put(excelConfig.getKeyColCarryCheName(),   cell.getColumnIndex()); break;
                case "Put CHE Name":    encontradas.put(excelConfig.getKeyColPutCheName(),     cell.getColumnIndex()); break;
                case "From Position":   encontradas.put(excelConfig.getKeyColFromPosition(),   cell.getColumnIndex()); break;
                case "To Position":     encontradas.put(excelConfig.getKeyColToPosition(),     cell.getColumnIndex()); break;
                case "Carrier Visit":   encontradas.put(excelConfig.getKeyColCarrierVisit(),   cell.getColumnIndex()); break;
            }
        }

        // Validar que todas las columnas requeridas estén presentes
        validarColumna(encontradas, excelConfig.getKeyColTimeCompleted());
        validarColumna(encontradas, excelConfig.getKeyColCraneCheName());
        validarColumna(encontradas, excelConfig.getKeyColFetchCheName());
        validarColumna(encontradas, excelConfig.getKeyColCarryCheName());
        validarColumna(encontradas, excelConfig.getKeyColPutCheName());
        validarColumna(encontradas, excelConfig.getKeyColFromPosition());
        validarColumna(encontradas, excelConfig.getKeyColToPosition());
        validarColumna(encontradas, excelConfig.getKeyColCarrierVisit());

        // Asignar índices detectados al modelo de configuración
        excelConfig.setColTimeCompleted(encontradas.get(excelConfig.getKeyColTimeCompleted()));
        excelConfig.setColCraneCheName( encontradas.get(excelConfig.getKeyColCraneCheName()));
        excelConfig.setColFetchCheName( encontradas.get(excelConfig.getKeyColFetchCheName()));
        excelConfig.setColCarryCheName( encontradas.get(excelConfig.getKeyColCarryCheName()));
        excelConfig.setColPutCheName(   encontradas.get(excelConfig.getKeyColPutCheName()));
        excelConfig.setColFromPosition( encontradas.get(excelConfig.getKeyColFromPosition()));
        excelConfig.setColToPosition(   encontradas.get(excelConfig.getKeyColToPosition()));
        excelConfig.setColCarrierVisit( encontradas.get(excelConfig.getKeyColCarrierVisit()));

        System.out.println("[OK] Columnas detectadas correctamente en MoveHistory.");
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
     * Recorre todas las filas de datos y las indexa en el caché.
     *
     * @param filtro Si es null, carga todas las visitas.
     *               Si tiene valores, solo carga las visitas contenidas en él.
     */
    private void cargarFilas(XSSFSheet hoja, Set<String> filtro) {
        int totalCargadas = 0;
        // Fila 0 = encabezado, datos desde fila 1
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            XSSFRow row = hoja.getRow(i);
            if (row == null) continue;

            String carrierVisit = getCellString(row.getCell(excelConfig.getColCarrierVisit())).trim();
            if (carrierVisit.isEmpty()) continue;
            if (filtro != null && !filtro.contains(carrierVisit)) continue;

            String timeCompleted = getCellString(row.getCell(excelConfig.getColTimeCompleted()));
            String craneCheName  = getCellString(row.getCell(excelConfig.getColCraneCheName()));
            String fetchCheName  = getCellString(row.getCell(excelConfig.getColFetchCheName()));
            String carryCheName  = getCellString(row.getCell(excelConfig.getColCarryCheName()));
            String putCheName    = getCellString(row.getCell(excelConfig.getColPutCheName()));
            String fromPosition  = getCellString(row.getCell(excelConfig.getColFromPosition()));
            String toPosition    = getCellString(row.getCell(excelConfig.getColToPosition()));

            ExcelMoveHistoryTeuCost.MoveHistoryTeuCost registro =
                    new ExcelMoveHistoryTeuCost.MoveHistoryTeuCost(
                            fetchCheName, putCheName, carryCheName,
                            craneCheName, fromPosition, toPosition,
                            timeCompleted, carrierVisit
                    );

            // Obtener (o crear) el mapa grúa→registros para esta visita
            Map<String, List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost>> porGrua =
                    cachePorVisitaYGrua.computeIfAbsent(carrierVisit, k -> new HashMap<>());

            // Indexar por fetchCheName (grúa origen)
            if (!fetchCheName.isEmpty()) {
                porGrua.computeIfAbsent(fetchCheName, k -> new ArrayList<>()).add(registro);
            }
            // Indexar por putCheName (grúa destino) si es diferente
            if (!putCheName.isEmpty() && !putCheName.equals(fetchCheName)) {
                porGrua.computeIfAbsent(putCheName, k -> new ArrayList<>()).add(registro);
            }

            totalCargadas++;
        }
        System.out.println("[INFO] Filas cargadas en caché: " + totalCargadas);
    }

    // =========================================================
    //  CONSULTAS DE DURACIÓN — reciben nroVisita
    // =========================================================

    /**
     * Retorna la duración en horas de todas las grúas conocidas para una visita.
     */
    public Map<String, Double> extraerResumenDuracionGruas(String nroVisita) {
        Map<String, Double> resumen = new LinkedHashMap<>();
        String[] gruas = {
                "STS01", "STS02", "STS03",
                "GM01",  "GM02",
                "RTG01", "RTG02", "RTG03", "RTG04", "RTG05", "RTG06",
                "S-501", "S-502", "S-503", "S-504", "S-505", "S-506", "S-507"
        };
        for (String grua : gruas) {
            resumen.put(grua, extraerDuracionGrua(nroVisita, grua));
        }
        return resumen;
    }

    /**
     * Calcula la duración de una grúa específica en una visita.
     * STS y GM → horas brutas (inicio a fin).
     * RTG y equipos menores → horas efectivas (excluye breaks > umbral).
     */
    public double extraerDuracionGrua(String nroVisita, String nombreGrua) {
        List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost> registros =
                obtenerRegistrosOrdenados(nroVisita, nombreGrua);

        if (registros.size() < 2) return 0.0;

        try {
            List<LocalDateTime> fechas = new ArrayList<>();
            for (ExcelMoveHistoryTeuCost.MoveHistoryTeuCost r : registros) {
                try {
                    fechas.add(LocalDateTime.parse(r.getTimeCompleted().trim(), FORMATTER));
                } catch (Exception e) {
                    System.out.println("[WARN] Fecha no parseable: '" + r.getTimeCompleted() + "'");
                }
            }
            if (fechas.size() < 2) return 0.0;

            long minutosBrutos = ChronoUnit.MINUTES.between(
                    fechas.get(0), fechas.get(fechas.size() - 1));

            long minutosEfectivos  = 0;
            LocalDateTime inicioSeg = fechas.get(0);
            for (int i = 1; i < fechas.size(); i++) {
                long intervalo = ChronoUnit.MINUTES.between(fechas.get(i - 1), fechas.get(i));
                if (intervalo > UMBRAL_BREAK_MINUTOS) {
                    minutosEfectivos += ChronoUnit.MINUTES.between(inicioSeg, fechas.get(i - 1));
                    inicioSeg = fechas.get(i);
                }
            }
            minutosEfectivos += ChronoUnit.MINUTES.between(inicioSeg, fechas.get(fechas.size() - 1));

            if (nombreGrua.startsWith("STS") || nombreGrua.startsWith("GM")) {
                return minutosBrutos / 60.0;
            } else {
                return minutosEfectivos / 60.0;
            }

        } catch (Exception e) {
            System.out.println("[ERROR] extraerDuracionGrua(" + nroVisita + "," + nombreGrua + "): "
                    + e.getMessage());
            return 0.0;
        }
    }

    // =========================================================
    //  CONSULTAS DE MOVIMIENTOS — reciben nroVisita
    // =========================================================

    public int extraerMovimientosRTG(String nroVisita) {
        return extraerMovimientosPorPrefijo(nroVisita, "RTG")
                .entrySet().stream()
                .filter(e -> e.getKey().matches("RTG0[1-4]"))
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int extraerMovimientosERTG(String nroVisita) {
        return extraerMovimientosPorPrefijo(nroVisita, "RTG")
                .entrySet().stream()
                .filter(e -> e.getKey().equals("RTG05") || e.getKey().equals("RTG06"))
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int extraerMovimientosSTS(String nroVisita) {
        return extraerMovimientosPorPrefijo(nroVisita, "STS")
                .entrySet().stream()
                .filter(e -> !e.getKey().startsWith("TOTAL_"))
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int extraerMovimientosStacker(String nroVisita) {
        return Arrays.asList("S-501", "S-502", "S-503", "S-506").stream()
                .mapToInt(k -> extraerMovimientosGrua(nroVisita, k).get("totalMovimientos"))
                .sum();
    }

    public int extraerMovimientosEmptyHand(String nroVisita) {
        return Arrays.asList("S-504", "S-505", "S-507").stream()
                .mapToInt(k -> extraerMovimientosGrua(nroVisita, k).get("totalMovimientos"))
                .sum();
    }

    public Map<String, Integer> extraerMovimientosGrua(String nroVisita, String nombreGrua) {
        Map<String, Integer> resumen = new HashMap<>();
        resumen.put("totalMovimientos", 0);
        resumen.put("comoFetch", 0);
        resumen.put("comoPut",   0);

        List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost> registros =
                obtenerRegistrosOrdenados(nroVisita, nombreGrua);
        if (registros.isEmpty()) return resumen;

        int fetch = 0, put = 0;
        for (ExcelMoveHistoryTeuCost.MoveHistoryTeuCost r : registros) {
            if (r.getFetCheName().equals(nombreGrua)) fetch++;
            if (r.getPutCheName().equals(nombreGrua)) put++;
        }
        resumen.put("totalMovimientos", fetch + put);
        resumen.put("comoFetch", fetch);
        resumen.put("comoPut",   put);
        return resumen;
    }

    public Map<String, Integer> extraerMovimientosPorPrefijo(String nroVisita, String prefijo) {
        Map<String, Integer> resumen = new LinkedHashMap<>();
        verificarCache();

        Map<String, List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost>> porGrua =
                cachePorVisitaYGrua.getOrDefault(nroVisita, Collections.emptyMap());

        porGrua.keySet().stream()
                .filter(k -> k.toUpperCase().startsWith(prefijo.toUpperCase()))
                .sorted()
                .forEach(grua -> {
                    List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost> filas =
                            obtenerRegistrosOrdenados(nroVisita, grua);
                    int fetch = 0, put = 0;
                    for (ExcelMoveHistoryTeuCost.MoveHistoryTeuCost r : filas) {
                        if (r.getFetCheName().equals(grua)) fetch++;
                        if (r.getPutCheName().equals(grua)) put++;
                    }
                    resumen.put(grua, fetch + put);
                });

        int total = resumen.values().stream().mapToInt(Integer::intValue).sum();
        resumen.put("TOTAL_" + prefijo.toUpperCase(), total);
        return resumen;
    }

    // =========================================================
    //  UTILIDADES PRIVADAS
    // =========================================================

    private List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost> obtenerRegistrosOrdenados(
            String nroVisita, String nombreGrua) {
        verificarCache();

        Map<String, List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost>> porGrua =
                cachePorVisitaYGrua.getOrDefault(nroVisita, Collections.emptyMap());

        List<ExcelMoveHistoryTeuCost.MoveHistoryTeuCost> registros = new ArrayList<>(
                porGrua.getOrDefault(nombreGrua, Collections.emptyList()));

        registros.sort((a, b) -> {
            try {
                return LocalDateTime.parse(a.getTimeCompleted().trim(), FORMATTER)
                        .compareTo(LocalDateTime.parse(b.getTimeCompleted().trim(), FORMATTER));
            } catch (Exception e) { return 0; }
        });
        return registros;
    }

    private void verificarCache() {
        if (cachePorVisitaYGrua == null) {
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
    //  SETTER
    // =========================================================

    public void setNombreHoja(String nombreHoja) { this.nombreHoja = nombreHoja; }
}