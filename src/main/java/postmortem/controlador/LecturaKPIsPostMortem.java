package postmortem.controlador;

import postmortem.modelo.ExcelMoveHistoryPostMortem;
import postmortem.modelo.ExcelNavesPostMortem;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class LecturaKPIsPostMortem {

    private static final SimpleDateFormat FMT = new SimpleDateFormat("dd-MMM-yy HH:mm", Locale.ENGLISH);

    private List<ExcelMoveHistoryPostMortem.DatosMoveHistory> excelMoveHistory;
    private List<ExcelNavesPostMortem.DatosNave> excelNaves;

    public LecturaKPIsPostMortem(String rutaExcelNaves, String rutaExcelMoveHistory) {
        LecturaNaves lecturaNaves = new LecturaNaves(rutaExcelNaves);
        LecturaMoveHistoryPostMortem lecturaMoveHistory = new LecturaMoveHistoryPostMortem(rutaExcelMoveHistory);

        this.excelNaves       = lecturaNaves.extraerDatosNave();
        this.excelMoveHistory = lecturaMoveHistory.extraerDatosMoveHistory();
    }

    // ==========================================================
    // PUNTO DE ENTRADA
    // ==========================================================
    public List<Map<String, Object>> calcularKpis() {
        return extraerIndicadoresTabla(extraerVisitas());
    }

    public List<String> extraerVisitas() {
        List<String> visitas = new ArrayList<>();
        for (ExcelNavesPostMortem.DatosNave nave : excelNaves) {
            visitas.add(nave.getVisita());
        }
        return visitas;
    }

    // ==========================================================
    // TABLA DE INDICADORES POR VISITA
    // ==========================================================
    private List<Map<String, Object>> extraerIndicadoresTabla(List<String> visitas) {
        List<Map<String, Object>> indicadores = new ArrayList<>();

        for (String visita : visitas) {

            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada =
                    filtrarMoveHistoryPorVisita(visita);

            ExcelNavesPostMortem.DatosNave nave = filtrarNave(visita);

            double duracionBerth     = duracionEnHoras(nave.getBerthATA(), nave.getBerthATD());
            double duracionOperacion = duracionEnHoras(nave.getInicioOperaciones(), nave.getFinOperaciones());

            int totalMovimientos = dataFiltrada.size();
            int movsSTS = contarMovimientosSTS(dataFiltrada);

            double gmph = extraerGMPHPorVisita(dataFiltrada);
            double bmph  = extraerBMPHPorVisita(totalMovimientos, duracionBerth);
            double craneIntensity = extraerCraneIntensity(dataFiltrada, duracionOperacion);

            Map<String, Integer> detalleMovimientos = extraerDetalleMovmientos(dataFiltrada);

            List<Map<String, Object>> tablaGruas = extraerTablaGruas(dataFiltrada);

            List<Map<String, Object>> demoras = new ArrayList<>();

            Map<String, Object> tabla = new HashMap<>();
            tabla.put("Visita", visita);
            tabla.put("Linea", nave.getLinea());
            tabla.put("Servicio", nave.getServicio());
            tabla.put("BerthATA", nave.getBerthATA());
            tabla.put("BerthATD", nave.getBerthATD());
            tabla.put("StartWork", nave.getInicioOperaciones());
            tabla.put("EndWork", nave.getFinOperaciones());
            tabla.put("DuracionBerth", duracionBerth);
            tabla.put("DuracionOp", duracionOperacion);
            tabla.put("TotalMovimientos", totalMovimientos);
            tabla.put("MovsSTS", movsSTS);
            tabla.put("GMPH", gmph);
            tabla.put("BMPH", bmph);
            tabla.put("CraneIntensity", craneIntensity);
            tabla.put("TablaGruas", tablaGruas);
            tabla.put("Descarga", detalleMovimientos.get("Descarga"));
            tabla.put("Embarque", detalleMovimientos.get("Embarque"));
            tabla.put("Restibas", detalleMovimientos.get("Restibas"));
            tabla.put("Transbordos", detalleMovimientos.get("Transbordos"));
            tabla.put("Demoras", demoras);

            // TODO : Para el detalle de los movimientos deberíamos de tomar la tabla de Chargeable Units Events

            indicadores.add(tabla);

            System.out.printf("[KPI] Visita=%-8s | Linea=%s | Servicio=%s | Movs=%3d | MovsSTS=%3d | "
                            + "DurBerth=%.2f hrs | DurOp=%.2f hrs | "
                            + "GMPH=%.2f | BMPH=%.2f | CraneInt=%.2f%n",
                    visita, nave.getLinea(), nave.getServicio(), totalMovimientos, movsSTS,
                    duracionBerth, duracionOperacion,
                    gmph, bmph, craneIntensity);
        }

        return indicadores;
    }


    // ==========================================================
    // GMPH — Gross Moves Per Hour por grúa
    //  Solo filas donde CraneCheName empieza con "STS"
    //  Duración = max(timeCompleted) - min(timeCompleted) por grúa
    //  GMPH total = sum(movs por grúa) / sum(hrs por grúa)
    // ==========================================================
    private Double extraerGMPHPorVisita(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada) {

        // Agrupar por grúa STS
        Map<String, List<ExcelMoveHistoryPostMortem.DatosMoveHistory>> porGrua =
                dataFiltrada.stream()
                        .filter(d -> esSTS(d.getCraneCheName()))
                        .collect(Collectors.groupingBy(
                                ExcelMoveHistoryPostMortem.DatosMoveHistory::getCraneCheName));

        if (porGrua.isEmpty()) return 0.0;

        double totalMovs = 0;
        double totalHrs  = 0;

        for (Map.Entry<String, List<ExcelMoveHistoryPostMortem.DatosMoveHistory>> entry
                : porGrua.entrySet()) {

            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> movsGrua = entry.getValue();
            double hrsGrua = duracionMinMaxTimeCompleted(movsGrua);

            if (hrsGrua > 0) {
                totalMovs += movsGrua.size();
                totalHrs  += hrsGrua;
            }
        }

        return totalHrs > 0 ? totalMovs / totalHrs : 0.0;
    }

    // ==========================================================
    // BMPH — Berth Moves Per Hour
    //  Total movimientos / duración entre BerthATA y BerthATD
    // ==========================================================
    private Double extraerBMPHPorVisita(int totalMovimientos, double duracionBerthHrs) {
        return duracionBerthHrs > 0 ? totalMovimientos / duracionBerthHrs : 0.0;
    }

    // ==========================================================
    // CRANE INTENSITY — Promedio de grúas STS activas por hora
    //  = sum(hrs individuales de cada grúa STS) / duracion operacion total
    // ==========================================================
    private Double extraerCraneIntensity(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada,
            double duracionOperacionHrs) {

        Map<String, List<ExcelMoveHistoryPostMortem.DatosMoveHistory>> porGrua =
                dataFiltrada.stream()
                        .filter(d -> esSTS(d.getCraneCheName()))
                        .collect(Collectors.groupingBy(
                                ExcelMoveHistoryPostMortem.DatosMoveHistory::getCraneCheName));

        if (porGrua.isEmpty() || duracionOperacionHrs <= 0) return 0.0;

        double sumaHrsGruas = 0;
        for (List<ExcelMoveHistoryPostMortem.DatosMoveHistory> movs : porGrua.values()) {
            sumaHrsGruas += duracionMinMaxTimeCompleted(movs);
        }

        return sumaHrsGruas / duracionOperacionHrs;
    }


    // ==========================================================
    // TABLA DE GRÚAS STS
    //  Por cada grúa: InicioOp, FinOp, Hrs, Movs, GMPH individual
    // ==========================================================
    private List<Map<String, Object>> extraerTablaGruas(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada) {

        Map<String, List<ExcelMoveHistoryPostMortem.DatosMoveHistory>> porGrua =
                dataFiltrada.stream()
                        .filter(d -> esSTS(d.getCraneCheName()))
                        .collect(Collectors.groupingBy(
                                ExcelMoveHistoryPostMortem.DatosMoveHistory::getCraneCheName));

        List<Map<String, Object>> tablaGruas = new ArrayList<>();

        for (Map.Entry<String, List<ExcelMoveHistoryPostMortem.DatosMoveHistory>> entry
                : porGrua.entrySet()) {

            String grua = entry.getKey();
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> movs = entry.getValue();

            List<Date> fechas = movs.stream()
                    .map(d -> parsearFecha(d.getTimeCompleted()))
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());

            if (fechas.isEmpty()) continue;

            Date inicioOp = Collections.min(fechas);
            Date finOp    = Collections.max(fechas);
            double hrs    = (finOp.getTime() - inicioOp.getTime()) / 3_600_000.0;
            int movsGrua  = movs.size();
            double gmphGrua = hrs > 0 ? movsGrua / hrs : 0.0;

            Map<String, Object> filaGrua = new HashMap<>();
            filaGrua.put("Grua", grua);
            filaGrua.put("InicioOp", FMT.format(inicioOp));
            filaGrua.put("FinOp", FMT.format(finOp));
            filaGrua.put("Hrs",  Math.round(hrs));
            filaGrua.put("Movs", movsGrua);
            filaGrua.put("GMPH", Math.round(gmphGrua * 100.0) / 100.0);

            tablaGruas.add(filaGrua);
        }

        tablaGruas.sort(Comparator.comparing(m -> (String) m.get("Grua")));
        return tablaGruas;
    }

    // ==========================================================
    // UTILIDADES
    // ==========================================================

    /** Duración en horas entre min y max timeCompleted de una lista de movimientos. */
    private double duracionMinMaxTimeCompleted(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> movs) {

        List<Date> fechas = movs.stream()
                .map(d -> parsearFecha(d.getTimeCompleted()))
                .filter(Objects::nonNull)
                .collect(Collectors.toList());

        if (fechas.size() < 2) return 0.0;
        Date min = Collections.min(fechas);
        Date max = Collections.max(fechas);
        return (max.getTime() - min.getTime()) / 3_600_000.0;
    }

    /** Duración en horas entre dos strings con formato "dd-MMM-yy HH:mm". */
    private double duracionEnHoras(String desde, String hasta) {
        Date dDesde = parsearFecha(desde);
        Date dHasta = parsearFecha(hasta);
        if (dDesde == null || dHasta == null) return 0.0;
        return (dHasta.getTime() - dDesde.getTime()) / 3_600_000.0;
    }

    /** Parsea "dd-MMM-yy HH:mm" → Date. Devuelve null si falla. */
    private Date parsearFecha(String valor) {
        if (valor == null || valor.trim().isEmpty()) return null;
        try {
            return FMT.parse(valor.trim());
        } catch (Exception e) {
            System.out.println("[WARN] No se pudo parsear fecha: '" + valor + "'");
            return null;
        }
    }

    /** true si el nombre de la grúa empieza con "STS" (insensible a mayúsculas). */
    private boolean esSTS(String craneCheName) {
        return craneCheName != null
                && craneCheName.trim().toUpperCase(Locale.ROOT).startsWith("STS");
    }

    /** Cuenta movimientos cuya grúa es STS. */
    private int contarMovimientosSTS(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada) {
        return dataFiltrada.stream().filter(data -> data.getCraneCheName().startsWith("STS")).collect(Collectors.toList()).size();
    }

    private List<ExcelMoveHistoryPostMortem.DatosMoveHistory> filtrarMoveHistoryPorVisita(
            String visita) {
        return excelMoveHistory.stream()
                .filter(d -> d.getVisita().equals(visita))
                .collect(Collectors.toList());
    }

    private ExcelNavesPostMortem.DatosNave filtrarNave(String visita) {
        return excelNaves.stream()
                .filter(d -> d.getVisita().equals(visita))
                .findFirst()
                .orElseThrow(() -> new RuntimeException(
                        "No se encontró nave para visita: " + visita));
    }

    private Map<String, Integer> extraerDetalleMovmientos(
            List<ExcelMoveHistoryPostMortem.DatosMoveHistory> dataFiltrada
    ){
        Map<String, Integer> detalleMov = new HashMap<>();
        int movimientosImportacion =
                (int) dataFiltrada.stream().filter(d -> d.getUnitCategory().equals("Import")).count();

        int movimientosExportacion =
                (int) dataFiltrada.stream().filter(d->d.getUnitCategory().equals("Export")).count();

        int movimientosRestiba =
                (int) dataFiltrada.stream().filter(d->d.getUnitCategory().equals("Through")).count();

        int movimientosTransbordo =
                (int) dataFiltrada.stream().filter(d->d.getUnitCategory().equals("Transship")).count();

        detalleMov.put("Descarga", movimientosImportacion);
        detalleMov.put("Embarque", movimientosExportacion);
        detalleMov.put("Restiba", movimientosRestiba);
        detalleMov.put("Transbordos", movimientosTransbordo);

        return detalleMov;
    }
}