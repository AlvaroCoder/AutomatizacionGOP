package teucost.modelos;

import java.sql.SQLOutput;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.stream.Collector;
import java.util.stream.Collectors;

public class LecturaMoveHistory {
    private String rutaMoveHistoryResumen;
    private String nroVisita;
    private String nombreHojaMoveHistory = "MoveEvent";

    private final int colMovKind       = 2;
    private final int colTimeCompleted = 0;
    private final int colUnitCategory  = 3;
    private final int colUnitNbr       = 4;
    private final int colFetchCheName  = 6;
    private final int colCarryCheName  = 7;
    private final int colPutCheName    = 8;
    private final int colNroVisita     = 11;
    private final int filaInicioDatos  = 2;

    private static final DateTimeFormatter FORMATTER =
            DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);

    private Map<String, List<String[]>> cachePorRTG = null;

    public LecturaMoveHistory(String rutaMoveHistoryResumen) {
        this.rutaMoveHistoryResumen = rutaMoveHistoryResumen;
    }

    public void cargarDatos() {
        cachePorRTG = new HashMap<>();
        LecturaExcels lecturaExcels = null;
        try {
            lecturaExcels = new LecturaExcels(this.rutaMoveHistoryResumen);
            List<String> colVisita = lecturaExcels.leerColumna(nombreHojaMoveHistory, colNroVisita);
            for (int i = filaInicioDatos; i < colVisita.size(); i++) {
                String visita = colVisita.get(i).trim();
                if (visita.isEmpty() || !visita.equals(this.nroVisita)) continue;

                String timeCompleted = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colTimeCompleted);
                String movKind       = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colMovKind);
                String unitNBR       = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colUnitNbr);
                String unitCategory  = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colUnitCategory);
                String fetchCheName  = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colFetchCheName);
                String putCheName    = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colPutCheName);
                String carryCheName  = lecturaExcels.leerCelda(nombreHojaMoveHistory, i, colCarryCheName);
                String[] fila = {timeCompleted, movKind, unitCategory, fetchCheName, putCheName, visita, unitNBR, carryCheName};
                // Indexar por fetchCheName y putCheName para acceso O(1) por RTG
                cachePorRTG.computeIfAbsent(fetchCheName, k -> new ArrayList<>()).add(fila);
                if (!putCheName.equals(fetchCheName)) {
                    cachePorRTG.computeIfAbsent(putCheName, k -> new ArrayList<>()).add(fila);
                }
            }

            System.out.println("[INFO] Datos cargados — RTGs encontrados: " + cachePorRTG.keySet());

        } catch (Exception err) {
            System.out.println("[ERROR] cargarDatos: " + err.getMessage());
            err.printStackTrace();
        } finally {
            if (lecturaExcels != null) {
                try { lecturaExcels.cerrar(); } catch (Exception ignored) {}
            }
        }
    }

    private List<String[]> obtenerFilasRTGOrdenadas(String nombreRTG) {
        if (cachePorRTG == null) {
            System.out.println("[WARN] Cache vacío — llama a cargarDatos() primero.");
            return Collections.emptyList();
        }

        List<String[]> filas = new ArrayList<>(
                cachePorRTG.getOrDefault(nombreRTG, Collections.emptyList()));

        filas.sort((a, b) -> {
            try {
                return LocalDateTime.parse(a[0].trim(), FORMATTER)
                        .compareTo(LocalDateTime.parse(b[0].trim(), FORMATTER));
            } catch (Exception e) { return 0; }
        });
        return filas;
    }

    public double extraerDuracionGrua(String nombreGrua) {
        List<String[]> filas = obtenerFilasRTGOrdenadas(nombreGrua);
        if (filas.size() < 2) {
            return 0.0;
        }

        try {
            final int UMBRAL_BREAK_MINUTOS = 15;

            long minutosEfectivos = 0;
            int  cantidadBreaks   = 0;

            List<LocalDateTime> fechas = new ArrayList<>();
            for (String[] fila : filas) {
                try {
                    fechas.add(LocalDateTime.parse(fila[0].trim(), FORMATTER));
                } catch (Exception e) {
                    System.out.println("[WARN] No se pudo parsear fecha: '" + fila[0] + "'");
                }
            }

            LocalDateTime inicioSegmento = fechas.get(0);

            for (int i = 1; i < fechas.size(); i++) {
                LocalDateTime anterior = fechas.get(i - 1);
                LocalDateTime actual   = fechas.get(i);
                long intervalo = ChronoUnit.MINUTES.between(anterior, actual);

                if (intervalo > UMBRAL_BREAK_MINUTOS) {
                    long minutosSegmento = ChronoUnit.MINUTES.between(inicioSegmento, anterior);
                    minutosEfectivos += minutosSegmento;
                    cantidadBreaks++;
                    inicioSegmento = actual;
                }
            }

            long minutosUltimoSegmento = ChronoUnit.MINUTES.between(
                    inicioSegmento, fechas.get(fechas.size() - 1));
            minutosEfectivos += minutosUltimoSegmento;

            // ✅ Calcular minutos totales brutos
            long minutosBrutos = ChronoUnit.MINUTES.between(
                    fechas.get(0), fechas.get(fechas.size() - 1));

            double horasEfectivas = minutosEfectivos / 60.0;
            double horasBrutas    = minutosBrutos    / 60.0;
            System.out.println("horasBrutas = " + horasBrutas);
            if (nombreGrua.startsWith("STS") || nombreGrua.startsWith("GM")) {
                return horasBrutas;
            } else {
                return horasEfectivas;
            }

        } catch (Exception e) {
            System.out.println("[ERROR] extraerDuracionGrua: " + e.getMessage());
            return 0.0;
        }
    }

    public Map<String, Integer> extraerMovimientosGrua(String nombreGrua) {
        Map<String, Integer> resumen = new HashMap<>();
        resumen.put("totalMovimientos", 0);
        resumen.put("comoFetch", 0); // grúa como origen  (DSCH)
        resumen.put("comoPut", 0); // grúa como destino (LOAD)

        List<String[]> filas = obtenerFilasRTGOrdenadas(nombreGrua);
        if (filas.isEmpty()) {
            System.out.println("[WARN] Sin movimientos para grúa: " + nombreGrua);
            return resumen;
        }

        int comoFetch = 0;
        int comoPut   = 0;

        for (String[] fila : filas) {
            String fetchChe = fila[3]; // columna fetchCheName
            String putChe   = fila[4]; // columna putCheName

            if (fetchChe.equals(nombreGrua)) comoFetch++;
            if (putChe.equals(nombreGrua))   comoPut++;
        }

        int total = comoFetch + comoPut;
        resumen.put("totalMovimientos", total);
        resumen.put("comoFetch",        comoFetch);
        resumen.put("comoPut",          comoPut);

        return resumen;
    }


    // ✅ Bug 1 corregido: usar extraerMovimientosPorPrefijo en lugar de grúas hardcodeadas
    public int extraerMovimientosRTG() {
        return extraerMovimientosPorPrefijo("RTG")
                .entrySet().stream()
                .filter(e -> e.getKey().equals("RTG01")
                        || e.getKey().equals("RTG02")
                        || e.getKey().equals("RTG03")
                        || e.getKey().equals("RTG04"))
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int extraerMovimientosSTS() {
        return extraerMovimientosPorPrefijo("STS")
                .entrySet().stream()
                .filter(e -> !e.getKey().startsWith("TOTAL_"))
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int extraerMovimientosERTG() {
        // RTG05 y RTG06 son eRTG — filtrar solo esos del prefijo RTG
        return cachePorRTG == null ? 0 :
                cachePorRTG.keySet().stream()
                        .filter(k -> k.equals("RTG05") || k.equals("RTG06"))
                        .mapToInt(k -> extraerMovimientosGrua(k).get("totalMovimientos"))
                        .sum();
    }

    public int extraerMovimientosStacker() {
        return Arrays.asList("S-501", "S-502", "S-503", "S-506")
                .stream()
                .mapToInt(k -> extraerMovimientosGrua(k).get("totalMovimientos"))
                .sum();
    }

    public int extraerMovimientosEmptyHand() {
        return Arrays.asList("S-504", "S-505", "S-507")
                .stream()
                .mapToInt(k -> extraerMovimientosGrua(k).get("totalMovimientos"))
                .sum();
    }


    public Map<String, Integer> extraerMovimientosPorPrefijo(String prefijo) {
        Map<String, Integer> resumen = new LinkedHashMap<>();
        int totalGeneral = 0;

        if (cachePorRTG == null) {
            System.out.println("[WARN] Cache vacío — llama a cargarDatos() primero.");
            return resumen;
        }

        // Filtrar solo las claves que empiecen con el prefijo indicado
        List<String> gruasFiltradas = cachePorRTG.keySet().stream()
                .filter(k -> k.toUpperCase().startsWith(prefijo.toUpperCase()))
                .sorted()
                .collect(Collectors.toList());

        if (gruasFiltradas.isEmpty()) {
            System.out.println("[WARN] No se encontraron grúas con prefijo: '" + prefijo + "'");
            return resumen;
        }

        System.out.println("==============================================");
        System.out.println(" Prefijo buscado : " + prefijo);
        System.out.println("----------------------------------------------");
        System.out.printf(" %-15s %-10s %-10s %-10s%n", "GRÚA", "FETCH", "PUT", "TOTAL");
        System.out.println("----------------------------------------------");

        for (String nombreGrua : gruasFiltradas) {
            List<String[]> filas = obtenerFilasRTGOrdenadas(nombreGrua);

            int comoFetch = 0;
            int comoPut   = 0;
            for (String[] fila : filas) {
                if (fila[3].equals(nombreGrua)) comoFetch++;
                if (fila[4].equals(nombreGrua)) comoPut++;
            }

            int totalGrua = comoFetch + comoPut;
            totalGeneral += totalGrua;

            resumen.put(nombreGrua, totalGrua);

            System.out.printf(" %-15s %-10d %-10d %-10d%n",
                    nombreGrua, comoFetch, comoPut, totalGrua);
        }

        System.out.println("----------------------------------------------");
        System.out.printf(" %-15s %-10s %-10s %-10d%n", "TOTAL", "", "", totalGeneral);
        System.out.println("==============================================");

        resumen.put("TOTAL_" + prefijo.toUpperCase(), totalGeneral);

        return resumen;
    }

    public Map<String, Double> extraerResumenDuracionGruas(){
        Map<String, Double> resumenDuracion = new HashMap<>();

        resumenDuracion.put("STS01", extraerDuracionGrua("STS01"));
        resumenDuracion.put("STS02", extraerDuracionGrua("STS02"));
        resumenDuracion.put("STS03", extraerDuracionGrua("STS03"));
        resumenDuracion.put("GM01", extraerDuracionGrua("GM01"));
        resumenDuracion.put("GM02", extraerDuracionGrua("GM02"));
        resumenDuracion.put("RTG01", extraerDuracionGrua("RTG01"));
        resumenDuracion.put("RTG02", extraerDuracionGrua("RTG02"));
        resumenDuracion.put("RTG03", extraerDuracionGrua("RTG03"));
        resumenDuracion.put("RTG04", extraerDuracionGrua("RTG04"));
        resumenDuracion.put("RTG05", extraerDuracionGrua("RTG05"));
        resumenDuracion.put("RTG06", extraerDuracionGrua("RTG06"));
        resumenDuracion.put("S-501", extraerDuracionGrua("S-501"));
        resumenDuracion.put("S-502", extraerDuracionGrua("S-502"));
        resumenDuracion.put("S-503", extraerDuracionGrua("S-503"));
        resumenDuracion.put("S-504", extraerDuracionGrua("S-504"));
        resumenDuracion.put("S-505", extraerDuracionGrua("S-505"));
        resumenDuracion.put("S-506", extraerDuracionGrua("S-506"));
        resumenDuracion.put("S-507", extraerDuracionGrua("S-507"));

        return resumenDuracion;
    }

    public void imprimirTablaPorVisita() {
        if (cachePorRTG == null) { cargarDatos(); }

        // Unificar todas las filas de todos los RTGs sin duplicados
        Set<String[]> setFilas = Collections.newSetFromMap(new IdentityHashMap<>());
        cachePorRTG.values().forEach(setFilas::addAll);
        List<String[]> todasLasFilas = new ArrayList<>(setFilas);
        todasLasFilas.sort((a, b) -> {
            try {
                return LocalDateTime.parse(a[0].trim(), FORMATTER)
                        .compareTo(LocalDateTime.parse(b[0].trim(), FORMATTER));
            } catch (Exception e) { return 0; }
        });

        System.out.println("==================== TABLA MOVIMIENTOS ====================");
        System.out.printf("%-22s %-15s %-15s %-12s %-15s %-15s %-15s%n",
                "TIME COMPLETED", "MOV KIND", "UNIT NBR", "UNIT CAT", "FETCH CHE", "PUT CHE", "VISITA");
        System.out.println("-----------------------------------------------------------");
        todasLasFilas.forEach(f -> System.out.printf("%-22s %-15s %-15s %-12s %-15s %-15s %-15s%n",
                f[0], f[1], f[6], f[2], f[3], f[4], f[5]));
        System.out.println("-----------------------------------------------------------");
        System.out.println(" Visita: " + nroVisita + " | Total movimientos: " + todasLasFilas.size());
        System.out.println("===========================================================");
    }

    public void setRutaMoveHistoryResumen(String r) { this.rutaMoveHistoryResumen = r; }
    public void setNroVisita(String nroVisita)      { this.nroVisita = nroVisita; }

    public void setNombreHojaMoveHistory(String nombreHojaMoveHistory) {this.nombreHojaMoveHistory = nombreHojaMoveHistory;}

}