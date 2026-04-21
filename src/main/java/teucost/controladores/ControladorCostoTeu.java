package teucost.controladores;


import teucost.modelos.*;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ControladorCostoTeu {

    private String nroVisita;

    private String rutaExcelResumen;
    private String rutaNavesMes;

    // Filas y columnas del Excel de Naves por mes de cada año
    private final int filaCuadrillas     = 24;
    private final int colCuadrillas      = 2;
    private int colResumenCuadrillas     = 10;
    private String nombreHojaExcelResumen = "Resumen Visitas Feb 26";
    private String nombreHojaCosmos     = "";
    private int colVisitaExcelResumen    = 1;
    private int numMesOM                = 0;
    private String nombreHojaMoveHistory = "MoveEvent";

    private double costoDepreciacionSTS = 56.14;
    private double costoDepreciacionGM = 45.54;
    private double costoDepreciacionERTG = 15.42;
    private double costoDepreciacionRTG = 9.21;
    private double costoDepreciacionITV = 0.65;
    private double costoDepreciacionRSK = 2.14;
    private double costoDepreciacionEH = 1.49;

    private double promedioVueltasITV = 7;

    // Bancos de datos
    private String rutaThroughput;
    private String rutaControlCosmos;
    private String rutaConciliado;
    private String rutaOMMensual;
    private String rutaMoveHistory;

    // Excel Destino
    private String rutaExcelDestino;

    // Índice de hoja destino (0-based)
    private final int hojaDestinoCostos = 0;
    private final int hojaDestinoNave = 1;

    // =========================================================
    //  COLUMNAS DE LA TABLA DESTINO (0-based)
    //  MES | NOMBRE NAVE | VISITA | LINEA | FECHA FIN | SEMANA
    //  | CUADRILLAS | MOV CONT | TEUS | MUELLE | TIEMPO EFECT
    //  | COSTO ESTIBA | COSTO RTG
    // =========================================================
    private final int colMes            = 0;
    private final int colNombreNave     = 1;
    private final int colVisitaTabla    = 2;
    private final int colLinea          = 3;
    private final int colFechaFin       = 4;
    private final int colSemana         = 5;
    private final int colCuadrillasTabla= 6;
    private final int colMovCont        = 7;
    private final int colTeus           = 8;
    private final int colMuelle         = 9;
    private final int colTiempoEfectivo = 10;
    private final int colCostoEstiba    = 11;
    private final int colCostoRTG       = 12;
    private final int colCostoReachStack = 13;
    private final int colCostoEmptyHand = 14;
    private final int colCostoGruaMov   = 15;
    private final int colCostoITV       = 16;
    private final int colSubtotalCosto  = 17;
    private final int colCostoSTS01     = 18;
    private final int colCostoSTS02     = 19;
    private final int colCostoSTS03     = 20;
    private final int colCostoRTG05     = 21;
    private final int colCostoRTG06     = 22;
    private final int colCostoSubtotalSTSRTG = 23;
    private final int colCostoDeprecSTS01 = 24;
    private final int colCostoDeprecSTS02 = 25;
    private final int colCostoDeprecSTS03 = 26;
    private final int colCostoDeprecGM01 = 27;
    private final int colCostoDeprecGM02 = 28;
    private final int colCostoDeprecRTG05 = 29;
    private final int colCostoDeprecRTG06 = 30;
    private final int colCostoDeprecRTG01 = 31;
    private final int colCostoDeprecRTG02 = 32;
    private final int colCostoDeprecRTG03 = 33;
    private final int colCostoDeprecRTG04 = 34;
    private final int colCostoDeprecS501 = 35;
    private final int colCostoDeprecS502 = 36;
    private final int colCostoDeprecS503 = 37;
    private final int colCostoDeprecS506 = 38;
    private final int colCostoDeprecS504 = 39;
    private final int colCostoDeprecS505 = 40;
    private final int colCostoDeprecS507 = 41;
    private final int colCostoDeprecTT = 42;
    private final int colCostoDepreciacion = 43;
    private final int colCostoNaveOperativo = 44;
    private final int colCostoTeu = 45;

    // =========================================================
    //  COLUMNAS DE LA TABLA DESTINO (1-based)
    //  MES | NOMBRE NAVE | VISITA | LINEA | FECHA FIN | SEMANA
    //  | CUADRILLAS | MOV CONT | TEUS | MUELLE | TIEMPO EFECT
    //  | COSTO ESTIBA | COSTO RTG
    // =========================================================

    private EscritorDetalleNave escritorDetalle;

    // =========================================================
    //  CONSTRUCTORES
    // =========================================================

    public ControladorCostoTeu(
            String rutaThroughput,
            String rutaControlCosmos,
            String rutaConciliado,
            String rutaOMMensual,
            String rutaMoveHistory) {
        this.rutaThroughput    = rutaThroughput;
        this.rutaControlCosmos = rutaControlCosmos;
        this.rutaConciliado    = rutaConciliado;
        this.rutaOMMensual     = rutaOMMensual;
        this.rutaMoveHistory   = rutaMoveHistory;
    }

    public ControladorCostoTeu(String rutaExcelResumen) {
        this.rutaExcelResumen = rutaExcelResumen;
    }

    // =========================================================
    //  EXTRACCIÓN DE DATOS — un solo nroVisita
    // =========================================================

    /**
     * Extrae todos los datos necesarios para una visita y los encapsula
     * en un HashMap listo para escribir en la tabla destino.
     */
    public HashMap<String, Object> extraerDatosCostos(String nroVisita) {
        HashMap<String, Object> datos = new HashMap<>();
        try {
            System.out.println("\n[INFO] ═══ Procesando visita: " + nroVisita + " ═══");

            // --- Throughput: movimientos y TEUs ---
            LecturaThroughput lecturaThroughput = new LecturaThroughput(this.rutaThroughput);
            lecturaThroughput.setNombreHoja("Flash Report");
            lecturaThroughput.setNroVisita(nroVisita);
            HashMap<String, Integer> resumenMovTeus = lecturaThroughput.extraerResumenMovTeus();
            int sumaMovCont = resumenMovTeus.get("sumaMovContenedores");
            int sumaTeus    = resumenMovTeus.get("sumaTeus");

            // --- Control Cosmos: datos generales de la nave ---
            LecturaControlCosmos lecturaControlCosmos = new LecturaControlCosmos(this.rutaControlCosmos);
            lecturaControlCosmos.setNombreHojaCosmos(nombreHojaCosmos);
            lecturaControlCosmos.setNroVisita(nroVisita);
            HashMap<String, String> resumenCosmos = lecturaControlCosmos.extraerResumenNaveCuad();
            String nombreVisita     = resumenCosmos.getOrDefault("nombreVisita",     "");
            String lineaServicio    = resumenCosmos.getOrDefault("lineaServicio",    "");
            String cuadrilla        = resumenCosmos.getOrDefault("cuadrillas",       "0");
            String fechaFinalizacion= resumenCosmos.getOrDefault("fechaFinalizacion","");
            String nroSemana        = resumenCosmos.getOrDefault("semana",           "");
            String tiempoEfectivo   = resumenCosmos.getOrDefault("tiempoEfectivo",   "0");

            // --- Conciliado: muelle ---
            LecturaConciliado lecturaConciliado = new LecturaConciliado(this.rutaConciliado);
            String muelle = lecturaConciliado.extraerMuelleNave(nroVisita);

            // --- OM Mensual: costos ---
            LecturaOmMensual lecturaOmMensual = new LecturaOmMensual(this.rutaOMMensual);
            lecturaOmMensual.setNumMes(numMesOM > 0 ? numMesOM : extraerNumeroMesDesdeFecha(fechaFinalizacion));
            lecturaOmMensual.setNombreHojaOMDiesel("Fuel26");
            lecturaOmMensual.setNombreHojaEnergia("WS-Energy");
            lecturaOmMensual.setNombreHojaEstibas("Stevedoring26");

            double costoEstiba        = lecturaOmMensual.extraerCostoEstiba();
            double totalCostoEstiba   = costoEstiba * parsearDouble(cuadrilla);
            HashMap<String, Double> costosDiesel = lecturaOmMensual.extraerResumenCostosDiesel();
            double costoPorMovRTG     = costosDiesel.getOrDefault("costoPorMovRTG", 0.0);
            double costoPorMovStacker = costosDiesel.getOrDefault("costoPorMovSTK", 0.0);
            double costoPorMovEH      = costosDiesel.getOrDefault("costoPorMovEH", 0.0);
            double costoGMGalonHora   = costosDiesel.getOrDefault("costoGMGalonHora", 0.0);
            double costoTTHora        = costosDiesel.getOrDefault("costoPorMovTT", 0.0);
            double costoUnitGalon     = costosDiesel.getOrDefault("costoUnitGalon", 0.0);

            // Costo ITV

            // --- MoveHistory: movimientos RTG ---
            LecturaMoveHistory lecturaMoveHistory = new LecturaMoveHistory(this.rutaMoveHistory);
            lecturaMoveHistory.setNombreHojaMoveHistory(nombreHojaMoveHistory);
            lecturaMoveHistory.setNroVisita(nroVisita);
            lecturaMoveHistory.cargarDatos();

            int movimientosRTG  = lecturaMoveHistory.extraerMovimientosRTG();
            double costoTotalRTG = movimientosRTG * costoPorMovRTG;

            int movimientosReachStacker = lecturaMoveHistory.extraerMovimientosStacker();
            double costoTotalReachStacker = movimientosReachStacker * costoPorMovStacker;

            int movimientosEmptyHandler = lecturaMoveHistory.extraerMovimientosEmptyHand();
            double costoTotalEmptyHandler = movimientosEmptyHandler * costoPorMovEH;

            Map<String, Double> duracionGruas = lecturaMoveHistory.extraerResumenDuracionGruas();

            double duracionSTS01 = duracionGruas.getOrDefault("STS01", 0.0);
            double duracionSTS02 = duracionGruas.getOrDefault("STS02", 0.0);
            double duracionSTS03 = duracionGruas.getOrDefault("STS03", 0.0);
            double duracionGM01 = duracionGruas.getOrDefault("GM01", 0.0);
            double duracionGM02 = duracionGruas.getOrDefault("GM02", 0.0);
            double duracionRTG01 = duracionGruas.getOrDefault("RTG01", 0.0);
            double duracionRTG02 = duracionGruas.getOrDefault("RTG02", 0.0);
            double duracionRTG03 = duracionGruas.getOrDefault("RTG03", 0.0);
            double duracionRTG04 = duracionGruas.getOrDefault("RTG04", 0.0);
            double duracionRTG05 = duracionGruas.getOrDefault("RTG05", 0.0);
            double duracionRTG06 = duracionGruas.getOrDefault("RTG06", 0.0);
            double duracionS501 = duracionGruas.getOrDefault("S-501", 0.0);
            double duracionS502 = duracionGruas.getOrDefault("S-502", 0.0);
            double duracionS503 = duracionGruas.getOrDefault("S-503", 0.0);
            double duracionS504 = duracionGruas.getOrDefault("S-504", 0.0);
            double duracionS505 = duracionGruas.getOrDefault("S-505", 0.0);
            double duracionS506 = duracionGruas.getOrDefault("S-506", 0.0);
            double duracionS507 = duracionGruas.getOrDefault("S-507", 0.0);

            double duracionGruaMovil = duracionGM01 + duracionGM02;

            double costoTotalGruaMovil = duracionGruaMovil * costoGMGalonHora;

            long gruasGMActivas = duracionGruas.entrySet().stream()
                    .filter(e -> e.getKey().startsWith("GM") && e.getValue() > 0.0)
                    .count();

            long gruasSTSActivas = duracionGruas.entrySet().stream()
                    .filter(e -> e.getKey().startsWith("STS") && e.getValue() > 0.0)
                    .count();

            double galonesITVHora = parsearDouble(tiempoEfectivo) * costoTTHora;
            double ITVAsignados = gruasGMActivas * 4 + gruasSTSActivas * 5;
            double costoTotalITV = ITVAsignados * galonesITVHora * costoUnitGalon;

            HashMap<String, Double> costosElectricos = lecturaOmMensual.extraerResumenCostosEnergia();
            double costoSTS01 = costosElectricos.getOrDefault("costoEnergiaSTS01", 0.0);
            double costoSTS02 = costosElectricos.getOrDefault("costoEnergiaSTS02",0.0);
            double costoSTS03 = costosElectricos.getOrDefault("costoEnergiaSTS03",0.0);

            double costoTotalSTS01 = costoSTS01 * duracionSTS01;
            double costoTotalSTS02 = costoSTS02 * duracionSTS02;
            double costoTotalSTS03 = costoSTS03 * duracionSTS03;

            double costoERTG05 = costosElectricos.getOrDefault("costoEnergiaRTG05",0.0);
            double costoERTG06 = costosElectricos.getOrDefault("costoEnergiaRTG06",0.0);
            double costoTotalERTG05 = costoERTG05 * duracionGruas.getOrDefault("RTG05",0.0);
            double costoTotalERTG06 = costoERTG06 * duracionGruas.getOrDefault("RTG06", 0.0);

            double subtotalCostoSTSERTG =
                    costoTotalSTS01 + costoTotalSTS02 + costoTotalSTS03
                            + costoTotalERTG05 + costoTotalERTG06;

            // Depreciaciones
            double costoDepreciacionSTS01 = duracionSTS01 * costoDepreciacionSTS;
            double costoDepreciacionSTS02 = duracionSTS02 * costoDepreciacionSTS;
            double costoDepreciacionSTS03 = duracionSTS03 * costoDepreciacionSTS;
            double costoDepreciacionGM01  = duracionGM01 * costoDepreciacionGM;
            double costoDepreciacionGM02  = duracionGM02 * costoDepreciacionGM;
            double costoDepreciacionRTG01 = duracionRTG01 * costoDepreciacionRTG;
            double costoDepreciacionRTG02 = duracionRTG02 * costoDepreciacionRTG;
            double costoDepreciacionRTG03 = duracionRTG03 * costoDepreciacionRTG;
            double costoDepreciacionRTG04 = duracionRTG04 * costoDepreciacionRTG;
            double costoDepreciacionRTG05 = duracionRTG05 * costoDepreciacionERTG;
            double costoDepreciacionRTG06 = duracionRTG06 * costoDepreciacionERTG;
            double costoDepreciacionS501 = duracionS501 * costoDepreciacionRSK;
            double costoDepreciacionS502 = duracionS502 * costoDepreciacionRSK;
            double costoDepreciacionS503 = duracionS503 * costoDepreciacionRSK;
            double costoDepreciacionS506 = duracionS506 * costoDepreciacionRSK;
            double costoDepreciacionS504 = duracionS504 * costoDepreciacionEH;
            double costoDepreciacionS505 = duracionS505 * costoDepreciacionEH;
            double costoDepreciacionS507 = duracionS507 * costoDepreciacionEH;
            double costoDepreciacionTT = ITVAsignados * promedioVueltasITV * costoDepreciacionITV;

            double sumaCostoDepreciacion =
                    costoDepreciacionSTS01 + costoDepreciacionSTS02 + costoDepreciacionSTS03
                            + costoDepreciacionGM01 + costoDepreciacionGM02  // ← corregido
                            + costoDepreciacionRTG01 + costoDepreciacionRTG02 + costoDepreciacionRTG03 + costoDepreciacionRTG04
                            + costoDepreciacionRTG05 + costoDepreciacionRTG06
                            + costoDepreciacionS501 + costoDepreciacionS502 + costoDepreciacionS503 + costoDepreciacionS506
                            + costoDepreciacionS504 + costoDepreciacionS505 + costoDepreciacionS507
                            + costoDepreciacionTT;

            double subtotalCosto =
                    costoTotalITV +
                            costoTotalRTG +
                            costoTotalEmptyHandler +
                            costoTotalReachStacker +
                            costoTotalGruaMovil;

            double sumaCostoNaveOperativo = totalCostoEstiba + subtotalCosto + subtotalCostoSTSERTG + sumaCostoDepreciacion;
            double costoPorTeu = sumaTeus > 0 ? sumaCostoNaveOperativo / sumaTeus : 0.0;

            // --- Extraer mes desde fechaFinalizacion (ej: "28-Feb-26 17:43" → "FEB") ---
            String mes = extraerMesDesdeFecha(fechaFinalizacion);

            // --- Empaquetar todos los datos ---
            datos.put("mes",             mes);
            datos.put("nombreVisita",    nombreVisita);
            datos.put("nroVisita",       nroVisita);
            datos.put("lineaServicio",   lineaServicio);
            datos.put("fechaFin",        fechaFinalizacion);
            datos.put("semana",          nroSemana);
            datos.put("cuadrillas",      parsearDouble(cuadrilla));
            datos.put("movCont",         (double) sumaMovCont);
            datos.put("teus",            (double) sumaTeus);
            datos.put("muelle",          muelle);
            datos.put("tiempoEfectivo",  parsearDouble(tiempoEfectivo));
            datos.put("costoEstiba",     totalCostoEstiba);
            datos.put("costoRTG",        costoTotalRTG);
            datos.put("costoReactStacker", costoTotalReachStacker);
            datos.put("costoEmptyHand", costoTotalEmptyHandler);
            datos.put("costoTotalGruaMovil", costoTotalGruaMovil);
            datos.put("costoTotalITV",   costoTotalITV);
            datos.put("subtotalCosto",   subtotalCosto);
            datos.put("costoTotalSTS01", costoTotalSTS01);
            datos.put("costoTotalSTS02", costoTotalSTS02);
            datos.put("costoTotalSTS03", costoTotalSTS03);
            datos.put("costoTotalRTG05", costoTotalERTG05);
            datos.put("costoTotalRTG06", costoTotalERTG06);
            datos.put("costoSubtotalSTSRTG", subtotalCostoSTSERTG);
            datos.put("costoDepreciacionSTS01", costoDepreciacionSTS01);
            datos.put("costoDepreciacionSTS02", costoDepreciacionSTS02);
            datos.put("costoDepreciacionSTS03", costoDepreciacionSTS03);
            datos.put("costoDepreciacionGM01", costoDepreciacionGM01);
            datos.put("costoDepreciacionGM02", costoDepreciacionGM02);
            datos.put("costoDepreciacionRTG05", costoDepreciacionRTG05);
            datos.put("costoDepreciacionRTG06", costoDepreciacionRTG06);
            datos.put("costoDepreciacionRTG01", costoDepreciacionRTG01);
            datos.put("costoDepreciacionRTG02", costoDepreciacionRTG02);
            datos.put("costoDepreciacionRTG03", costoDepreciacionRTG03);
            datos.put("costoDepreciacionRTG04", costoDepreciacionRTG04);
            datos.put("costoDepreciacionS501", costoDepreciacionS501);
            datos.put("costoDepreciacionS502", costoDepreciacionS502);
            datos.put("costoDepreciacionS503", costoDepreciacionS503);
            datos.put("costoDepreciacionS506", costoDepreciacionS506);
            datos.put("costoDepreciacionS504", costoDepreciacionS504);
            datos.put("costoDepreciacionS505", costoDepreciacionS505);
            datos.put("costoDepreciacionS507", costoDepreciacionS507);
            datos.put("costoDepreciacionTT", costoDepreciacionTT);
            datos.put("subtotalCostoDepreciacion", sumaCostoDepreciacion);
            datos.put("costoNaveOperativo", sumaCostoNaveOperativo);
            datos.put("costoPorTeu", costoPorTeu);
            datos.put("duracionRTG01", duracionRTG01);
            datos.put("duracionRTG02", duracionRTG02);
            datos.put("duracionRTG03", duracionRTG03);
            datos.put("duracionRTG04", duracionRTG04);
            datos.put("duracionRTG05", duracionRTG05);
            datos.put("duracionRTG06", duracionRTG06);
            datos.put("duracionSTS01", duracionSTS01);
            datos.put("duracionSTS02", duracionSTS02);
            datos.put("duracionSTS03", duracionSTS03);
            datos.put("duracionS501", duracionS501);
            datos.put("duracionS502", duracionS502);
            datos.put("duracionS503", duracionS503);
            datos.put("duracionS504", duracionS504);
            datos.put("duracionS505", duracionS505);
            datos.put("duracionS506", duracionS506);
            datos.put("duracionS507", duracionS507);

            datos.put("movimientosSTS01", lecturaMoveHistory.extraerMovimientosPorPrefijo("STS01").get("TOTAL_STS01"));
            datos.put("movimientosSTS02", lecturaMoveHistory.extraerMovimientosPorPrefijo("STS02").get("TOTAL_STS02"));
            datos.put("movimientosSTS03", lecturaMoveHistory.extraerMovimientosPorPrefijo("STS03").get("TOTAL_STS03"));
            datos.put("movimientosRTG01", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG01").get("TOTAL_RTG01"));
            datos.put("movimientosRTG02", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG02").get("TOTAL_RTG02"));
            datos.put("movimientosRTG03", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG03").get("TOTAL_RTG03"));
            datos.put("movimientosRTG04", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG04").get("TOTAL_RTG04"));
            datos.put("movimientosRTG05", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG05").get("TOTAL_RTG05"));
            datos.put("movimientosRTG06", lecturaMoveHistory.extraerMovimientosPorPrefijo("RTG06").get("TOTAL_RTG06"));
            datos.put("movimientosS501", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-501").get("TOTAL_S-501"));
            datos.put("movimientosS502", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-502").get("TOTAL_S-502"));
            datos.put("movimientosS503", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-503").get("TOTAL_S-503"));
            datos.put("movimientosS504", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-504").get("TOTAL_S-504"));
            datos.put("movimientosS505", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-505").get("TOTAL_S-505"));
            datos.put("movimientosS506", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-506").get("TOTAL_S-506"));
            datos.put("movimientosS507", lecturaMoveHistory.extraerMovimientosPorPrefijo("S-507").get("TOTAL_S-507"));


            System.out.println("[OK] Datos extraídos para visita: " + nroVisita);

        } catch (Exception err) {
            System.out.println("[ERROR] extraerDatosVisita (" + nroVisita + "): " + err.getMessage());
            err.printStackTrace();
        }

        return datos;
    }

    private int extraerNumeroMesDesdeFecha(String fechaStr) {
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);
            LocalDateTime dt = LocalDateTime.parse(fechaStr.trim(), formatter);
            return dt.getMonthValue();
        } catch (Exception e) {
            System.out.println("[WARN] No se pudo extraer número de mes de: '" + fechaStr + "'");
            return -1;
        }
    }

    public HashMap<String, Object> extraerDatosNave(String nroVisita){
        HashMap<String, Object> datos = new HashMap<>();
        try {
            LecturaThroughput lecturaThroughput = new LecturaThroughput(this.rutaThroughput);
            lecturaThroughput.setNombreHoja("Flash Report");
            lecturaThroughput.setNroVisita(nroVisita);
            HashMap<String, Integer> resumenMovTeus = lecturaThroughput.extraerResumenMovTeus();
            int sumaMovCont = resumenMovTeus.get("sumaMovContenedores");
            int sumaTeus    = resumenMovTeus.get("sumaTeus");

            LecturaControlCosmos lecturaControlCosmos = new LecturaControlCosmos(this.rutaControlCosmos);
            lecturaControlCosmos.setNroVisita(nroVisita);
            HashMap<String, String> resumenCosmos = lecturaControlCosmos.extraerResumenNaveCuad();
            String nombreVisita     = resumenCosmos.getOrDefault("nombreVisita",     "");
            String lineaServicio    = resumenCosmos.getOrDefault("lineaServicio",    "");
            String cuadrilla        = resumenCosmos.getOrDefault("cuadrillas",       "0");
            String fechaFinalizacion= resumenCosmos.getOrDefault("fechaFinalizacion","");
            String fechaInicio      = resumenCosmos.getOrDefault("fechaInicio",      "");
            String nroSemana        = resumenCosmos.getOrDefault("semana",           "");
            String tiempoEfectivo   = resumenCosmos.getOrDefault("tiempoEfectivo",   "0");

            String mes = extraerMesDesdeFecha(fechaFinalizacion);

            LecturaConciliado lecturaConciliado = new LecturaConciliado(this.rutaConciliado);
            String muelle = lecturaConciliado.extraerMuelleNave(nroVisita);

            LecturaMoveHistory lecturaMoveHistory = new LecturaMoveHistory(this.rutaMoveHistory);
            lecturaMoveHistory.setNroVisita(nroVisita);
            lecturaMoveHistory.cargarDatos();

            double movimientosSTS = lecturaMoveHistory.extraerMovimientosSTS();
            double movimientosRTG = lecturaMoveHistory.extraerMovimientosRTG();
            double movimientosERTG = lecturaMoveHistory.extraerMovimientosERTG();
            double movimientosRSK = lecturaMoveHistory.extraerMovimientosStacker();
            double movimientosEH  = lecturaMoveHistory.extraerMovimientosEmptyHand();

            datos.put("mes",                mes);
            datos.put("semana",             nroSemana);
            datos.put("nave",               nombreVisita);
            datos.put("visita",             nroVisita);
            datos.put("servicio",           lineaServicio);
            datos.put("muelle",             muelle);
            datos.put("fechaInicio",        fechaInicio);
            datos.put("fechaTermino",       fechaFinalizacion);
            datos.put("tiempoEfectivo",     tiempoEfectivo);
            datos.put("cuadrillas",         cuadrilla);
            datos.put("movimientosCont",    (double) sumaMovCont);
            datos.put("movimientosTEUS",    (double) sumaTeus);
            datos.put("movimientosSTS",     movimientosSTS);
            datos.put("movimientosRTG",     movimientosRTG);
            datos.put("movimientosERTG",    movimientosERTG);
            datos.put("movimientosStacker", movimientosRSK);
            datos.put("movimientosEmtpy", movimientosEH);

        } catch (Exception err){
            System.out.println("err = " + err.getMessage());
        }

        return datos;
    }

    // =========================================================
    //  ESCRITURA EN TABLA DESTINO — una visita
    // =========================================================

    /**
     * Escribe los datos de UNA visita en la tabla del Excel destino,
     * expandiendo la tabla y preservando el formato.
     */
    public void escribirVisitaEnTabla(String nroVisita) {
        HashMap<String, Object> datos = extraerDatosCostos(nroVisita);
        if (datos.isEmpty()) {
            System.out.println("[ERROR] No se obtuvieron datos para: " + nroVisita);
            return;
        }
        escribirFilaEnDestino(datos);          // hoja 0 — sin cambios
        escritorDetalle.escribirFila(datos);   // hoja 1 — delegado
    }

    /**
     * Escribe los datos de MÚLTIPLES visitas en la tabla destino,
     * una fila por visita, en el orden del array recibido.
     *
     * @param nrosVisita Lista de visitas a procesar. Ej: ["058-26", "073-26", "116-26"]
     */
    public void escribirVariasVisitasEnTabla(List<String> nrosVisita) {
        System.out.println("\n[INFO] ═══ Procesando lote de " + nrosVisita.size() + " visitas ═══");

        int exitosas = 0;
        int fallidas = 0;
        List<String> visitasFallidas = new ArrayList<>();

        for (String visita : nrosVisita) {
            try {
                escribirVisitaEnTabla(visita.trim());
                exitosas++;
            } catch (Exception e) {
                System.out.println("[ERROR] Falló visita '" + visita + "': " + e.getMessage());
                visitasFallidas.add(visita);
                fallidas++;
            }
        }

        // Resumen del lote
        System.out.println("\n════════ RESUMEN DEL LOTE ════════");
        System.out.println(" Total procesadas : " + nrosVisita.size());
        System.out.println(" Exitosas         : " + exitosas);
        System.out.println(" Fallidas         : " + fallidas);
        if (!visitasFallidas.isEmpty()) {
            System.out.println(" Visitas fallidas : " + visitasFallidas);
        }
        System.out.println("══════════════════════════════════");
    }

    // =========================================================
    //  ESCRITURA FÍSICA EN EXCEL — usa LecturaExcels
    // =========================================================

    private void escribirFilaEnDestino(HashMap<String, Object> datos) {
        LecturaExcels excelDestino = null;
        try {
            excelDestino = new LecturaExcels(this.rutaExcelDestino);

            // Insertar nueva fila en la tabla expandiendo su rango y copiando estilos
            Row nuevaFila = excelDestino.insertarFilaEnTabla(hojaDestinoCostos, 0);

            // --- Escribir cada columna ---
            setCeldaTexto(nuevaFila,   colMes,             getString(datos, "mes"));
            setCeldaTexto(nuevaFila,   colNombreNave,      getString(datos, "nombreVisita"));
            setCeldaTexto(nuevaFila,   colVisitaTabla,     getString(datos, "nroVisita"));
            setCeldaTexto(nuevaFila,   colLinea,           getString(datos, "lineaServicio"));
            setCeldaTexto(nuevaFila,   colFechaFin,        getString(datos, "fechaFin"));
            setCeldaTexto(nuevaFila,   colSemana,          getString(datos, "semana"));
            setCeldaNumero(nuevaFila,  colCuadrillasTabla, getDouble(datos, "cuadrillas"));
            setCeldaNumero(nuevaFila,  colMovCont,         getDouble(datos, "movCont"));
            setCeldaNumero(nuevaFila,  colTeus,            getDouble(datos, "teus"));
            setCeldaTexto(nuevaFila,   colMuelle,          getString(datos, "muelle"));
            setCeldaNumero(nuevaFila,  colTiempoEfectivo,  getDouble(datos, "tiempoEfectivo"));
            setCeldaNumero(nuevaFila,  colCostoEstiba,     getDouble(datos, "costoEstiba"));
            setCeldaNumero(nuevaFila,  colCostoRTG,        getDouble(datos, "costoRTG"));
            setCeldaNumero(nuevaFila, colCostoReachStack,  getDouble(datos, "costoReactStacker"));
            setCeldaNumero(nuevaFila, colCostoEmptyHand,   getDouble(datos, "costoEmptyHand"));
            setCeldaNumero(nuevaFila, colCostoGruaMov,     getDouble(datos, "costoTotalGruaMovil"));
            setCeldaNumero(nuevaFila, colCostoITV,         getDouble(datos, "costoTotalITV"));
            setCeldaNumero(nuevaFila, colSubtotalCosto,    getDouble(datos, "subtotalCosto"));
            setCeldaNumero(nuevaFila, colCostoSTS01,       getDouble(datos, "costoTotalSTS01"));
            setCeldaNumero(nuevaFila, colCostoSTS02,       getDouble(datos, "costoTotalSTS02"));
            setCeldaNumero(nuevaFila, colCostoSTS03,       getDouble(datos, "costoTotalSTS03"));
            setCeldaNumero(nuevaFila, colCostoRTG05,       getDouble(datos, "costoTotalRTG05"));
            setCeldaNumero(nuevaFila, colCostoRTG06,       getDouble(datos, "costoTotalRTG06"));
            setCeldaNumero(nuevaFila, colCostoSubtotalSTSRTG, getDouble(datos, "costoSubtotalSTSRTG"));
            setCeldaNumero(nuevaFila, colCostoDeprecSTS01, getDouble(datos, "costoDepreciacionSTS01") );
            setCeldaNumero(nuevaFila, colCostoDeprecSTS02, getDouble(datos, "costoDepreciacionSTS02"));
            setCeldaNumero(nuevaFila, colCostoDeprecSTS03, getDouble(datos, "costoDepreciacionSTS03"));
            setCeldaNumero(nuevaFila, colCostoDeprecGM01, getDouble(datos, "costoDepreciacionGM01"));
            setCeldaNumero(nuevaFila, colCostoDeprecGM02, getDouble(datos, "costoDepreciacionGM02"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG05, getDouble(datos, "costoDepreciacionRTG05"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG06, getDouble(datos, "costoDepreciacionRTG06"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG01, getDouble(datos, "costoDepreciacionRTG01"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG02, getDouble(datos, "costoDepreciacionRTG02"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG03, getDouble(datos, "costoDepreciacionRTG03"));
            setCeldaNumero(nuevaFila, colCostoDeprecRTG04, getDouble(datos, "costoDepreciacionRTG04"));
            setCeldaNumero(nuevaFila, colCostoDeprecS501, getDouble(datos, "costoDepreciacionS501"));
            setCeldaNumero(nuevaFila, colCostoDeprecS502, getDouble(datos, "costoDepreciacionS502"));
            setCeldaNumero(nuevaFila, colCostoDeprecS503, getDouble(datos, "costoDepreciacionS503"));
            setCeldaNumero(nuevaFila, colCostoDeprecS506, getDouble(datos, "costoDepreciacionS506"));
            setCeldaNumero(nuevaFila, colCostoDeprecS504, getDouble(datos, "costoDepreciacionS504"));
            setCeldaNumero(nuevaFila, colCostoDeprecS505, getDouble(datos, "costoDepreciacionS505"));
            setCeldaNumero(nuevaFila, colCostoDeprecS507, getDouble(datos, "costoDepreciacionS507"));
            setCeldaNumero(nuevaFila, colCostoDeprecTT, getDouble(datos, "costoDepreciacionTT"));
            setCeldaNumero(nuevaFila, colCostoDepreciacion, getDouble(datos, "subtotalCostoDepreciacion"));
            setCeldaNumero(nuevaFila, colCostoNaveOperativo, getDouble(datos, "costoNaveOperativo"));
            setCeldaNumero(nuevaFila, colCostoTeu, getDouble(datos, "costoPorTeu"));

            excelDestino.guardar();
            System.out.println("[OK] Fila escrita en tabla destino para visita: "
                    + getString(datos, "nroVisita"));

        } catch (Exception err) {
            System.out.println("[ERROR] escribirFilaEnDestino: " + err.getMessage());
            err.printStackTrace();
        } finally {
            if (excelDestino != null) {
                try { excelDestino.cerrar(); } catch (Exception ignored) {}
            }
        }
    }


    public void imprimirReporteDetalladoNave(String nroVisita) {
        HashMap<String, Object> datos = extraerDatosCostos(nroVisita);
        if (datos.isEmpty()) {
            System.out.println("[ERROR] No se pudieron obtener datos para: " + nroVisita);
            return;
        }

        // Extraer valores para facilitar el reporte
        double sumaTeus    = getDouble(datos, "teus");
        double sumaMovCont = getDouble(datos, "movCont");
        double costoEstiba = getDouble(datos, "costoEstiba");
        double costoRTG    = getDouble(datos, "costoRTG");
        double costoRSK    = getDouble(datos, "costoReactStacker");
        double costoEH     = getDouble(datos, "costoEmptyHand");
        double costoGM     = getDouble(datos, "costoTotalGruaMovil");
        double costoITV    = getDouble(datos, "costoTotalITV");
        double subtotal    = getDouble(datos, "subtotalCosto");

        double costoSTS01  = getDouble(datos, "costoTotalSTS01");
        double costoSTS02  = getDouble(datos, "costoTotalSTS02");
        double costoSTS03  = getDouble(datos, "costoTotalSTS03");
        double costoRTG05  = getDouble(datos, "costoTotalRTG05");
        double costoRTG06  = getDouble(datos, "costoTotalRTG06");
        double subtotalSTS = getDouble(datos, "costoSubtotalSTSRTG");

        double depSTS01    = getDouble(datos, "costoDepreciacionSTS01");
        double depSTS02    = getDouble(datos, "costoDepreciacionSTS02");
        double depSTS03    = getDouble(datos, "costoDepreciacionSTS03");
        double depGM01     = getDouble(datos, "costoDepreciacionGM01");
        double depGM02     = getDouble(datos, "costoDepreciacionGM02");
        double depRTG05    = getDouble(datos, "costoDepreciacionRTG05");
        double depRTG06    = getDouble(datos, "costoDepreciacionRTG06");
        double depRTG01    = getDouble(datos, "costoDepreciacionRTG01");
        double depRTG02    = getDouble(datos, "costoDepreciacionRTG02");
        double depRTG03    = getDouble(datos, "costoDepreciacionRTG03");
        double depRTG04    = getDouble(datos, "costoDepreciacionRTG04");
        double depS501     = getDouble(datos, "costoDepreciacionS501");
        double depS502     = getDouble(datos, "costoDepreciacionS502");
        double depS503     = getDouble(datos, "costoDepreciacionS503");
        double depS506     = getDouble(datos, "costoDepreciacionS506");
        double depS504     = getDouble(datos, "costoDepreciacionS504");
        double depS505     = getDouble(datos, "costoDepreciacionS505");
        double depS507     = getDouble(datos, "costoDepreciacionS507");
        double depTT       = getDouble(datos, "costoDepreciacionTT");
        double subtotalDep = getDouble(datos, "subtotalCostoDepreciacion");

        double costoOperativo = getDouble(datos, "costoNaveOperativo");
        double costoPorTeu    = getDouble(datos, "costoPorTeu");

        double duracionRTG01 = getDouble(datos, "duracionRTG01");
        double duracionRTG02 = getDouble(datos, "duracionRTG02");
        double duracionRTG03 = getDouble(datos, "duracionRTG03");
        double duracionRTG04 = getDouble(datos, "duracionRTG04");
        double duracionRTG05 = getDouble(datos, "duracionRTG05");
        double duracionRTG06 = getDouble(datos, "duracionRTG06");

        double duracionSTS01 = getDouble(datos, "duracionSTS01");
        double duracionSTS02 = getDouble(datos, "duracionSTS02");
        double duracionSTS03 = getDouble(datos, "duracionSTS03");
        String sep1  = "═==========";
        String sep2  = "─------------";
        String sep3  = "·.............";

        System.out.println("\n" + sep1);
        System.out.printf("  REPORTE DETALLADO — VISITA: %-34s%n", getString(datos, "nroVisita"));
        System.out.println(sep1);

        // ── INFORMACIÓN GENERAL ──────────────────────────────────
        System.out.println("\n  📋 INFORMACIÓN GENERAL");
        System.out.println(sep2);
        System.out.printf("  %-22s : %s%n",  "Nombre Nave",      getString(datos, "nombreVisita"));
        System.out.printf("  %-22s : %s%n",  "Visita",           getString(datos, "nroVisita"));
        System.out.printf("  %-22s : %s%n",  "Mes",              getString(datos, "mes"));
        System.out.printf("  %-22s : %s%n",  "Semana",           getString(datos, "semana"));
        System.out.printf("  %-22s : %s%n",  "Línea Servicio",   getString(datos, "lineaServicio"));
        System.out.printf("  %-22s : %s%n",  "Muelle",           getString(datos, "muelle"));
        System.out.printf("  %-22s : %s%n",  "Fecha Término",    getString(datos, "fechaFin"));
        System.out.printf("  %-22s : %.2f hrs%n", "Tiempo Efectivo", getDouble(datos, "tiempoEfectivo"));
        System.out.printf("  %-22s : %.0f%n","Cuadrillas",       getDouble(datos, "cuadrillas"));

        // ── MOVIMIENTOS ───────────────────────────────────────────
        System.out.println("\n  📦 MOVIMIENTOS");
        System.out.println(sep2);
        System.out.printf("  %-22s : %,.0f%n", "Mov. Contenedores", sumaMovCont);
        System.out.printf("  %-22s : %,.0f%n", "TEUs",              sumaTeus);
        System.out.println(sep3);
        System.out.printf("  %-22s : %,.0f%n", "Ratio TEU/Cont",
                sumaMovCont > 0 ? sumaTeus / sumaMovCont : 0);

        System.out.println("\n  📦 DURACIÓN RTG");
        System.out.println(sep2);
        System.out.printf("  %-22s : %,.0f%n", "RTG01", duracionRTG01);
        System.out.printf("  %-22s : %,.0f%n", "RTG02",  duracionRTG02 );
        System.out.printf("  %-22s : %,.0f%n", "RTG03",  duracionRTG03 );
        System.out.printf("  %-22s : %,.0f%n", "RTG04",  duracionRTG04 );
        System.out.printf("  %-22s : %,.0f%n", "RTG05",  duracionRTG05 );
        System.out.printf("  %-22s : %,.0f%n", "RTG06",  duracionRTG06 );

        System.out.println(sep3);

        System.out.println("\n  📦 DURACIÓN STS");
        System.out.println(sep2);
        System.out.printf("  %-22s : %,.0f%n", "STS01", duracionSTS01);
        System.out.printf("  %-22s : %,.0f%n", "STS02",  duracionSTS02 );
        System.out.printf("  %-22s : %,.0f%n", "STS03",  duracionSTS03 );


        System.out.println(sep3);

        // ── COSTO ESTIBA ─────────────────────────────────────────
        System.out.println("\n  👷 COSTO ESTIBA (Cuadrillas)");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "Total Estiba",      costoEstiba);
        System.out.printf("  %-22s : $ %,.2f%n", "Costo / TEU",       sumaTeus    > 0 ? costoEstiba / sumaTeus    : 0);
        System.out.printf("  %-22s : $ %,.2f%n", "Costo / Cont",      sumaMovCont > 0 ? costoEstiba / sumaMovCont : 0);

        // ── COSTO DIESEL (RTG, RSK, EH, GM, ITV) ─────────────────
        System.out.println("\n  ⛽ COSTO DIESEL");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG (diesel)",      costoRTG);
        System.out.printf("  %-22s : $ %,.2f%n", "Reach Stacker",     costoRSK);
        System.out.printf("  %-22s : $ %,.2f%n", "Empty Handler",     costoEH);
        System.out.printf("  %-22s : $ %,.2f%n", "Grúa Móvil (GM)",   costoGM);
        System.out.printf("  %-22s : $ %,.2f%n", "ITV",               costoITV);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL DIESEL",   subtotal);

        // ── COSTO ENERGÍA (STS + eRTG) ───────────────────────────
        System.out.println("\n  ⚡ COSTO ENERGÍA ELÉCTRICA");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "STS01",             costoSTS01);
        System.out.printf("  %-22s : $ %,.2f%n", "STS02",             costoSTS02);
        System.out.printf("  %-22s : $ %,.2f%n", "STS03",             costoSTS03);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG05",            costoRTG05);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG06",            costoRTG06);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL ENERGÍA",  subtotalSTS);

        // ── DEPRECIACIONES ────────────────────────────────────────
        System.out.println("\n  📉 DEPRECIACIONES");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "STS01",             depSTS01);
        System.out.printf("  %-22s : $ %,.2f%n", "STS02",             depSTS02);
        System.out.printf("  %-22s : $ %,.2f%n", "STS03",             depSTS03);
        System.out.printf("  %-22s : $ %,.2f%n", "GM01",              depGM01);
        System.out.printf("  %-22s : $ %,.2f%n", "GM02",              depGM02);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG05",            depRTG05);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG06",            depRTG06);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG01",             depRTG01);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG02",             depRTG02);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG03",             depRTG03);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG04",             depRTG04);
        System.out.printf("  %-22s : $ %,.2f%n", "S-501",             depS501);
        System.out.printf("  %-22s : $ %,.2f%n", "S-502",             depS502);
        System.out.printf("  %-22s : $ %,.2f%n", "S-503",             depS503);
        System.out.printf("  %-22s : $ %,.2f%n", "S-506",             depS506);
        System.out.printf("  %-22s : $ %,.2f%n", "S-504",             depS504);
        System.out.printf("  %-22s : $ %,.2f%n", "S-505",             depS505);
        System.out.printf("  %-22s : $ %,.2f%n", "S-507",             depS507);
        System.out.printf("  %-22s : $ %,.2f%n", "TT (ITV)",          depTT);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL DEPREC.",  subtotalDep);

        // ── RESUMEN FINAL ─────────────────────────────────────────
        System.out.println("\n" + sep1);
        System.out.println("  💰 RESUMEN FINAL DE COSTOS");
        System.out.println(sep1);
        System.out.printf("  %-22s : $ %,.2f%n", "Estiba",            costoEstiba);
        System.out.printf("  %-22s : $ %,.2f%n", "Diesel",            subtotal);
        System.out.printf("  %-22s : $ %,.2f%n", "Energía Eléctrica", subtotalSTS);
        System.out.printf("  %-22s : $ %,.2f%n", "Depreciación",      subtotalDep);
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "COSTO NAVE TOTAL",  costoOperativo);
        System.out.println(sep1);
        System.out.printf("  %-22s : $ %,.2f%n", "★ COSTO POR TEU",   costoPorTeu);
        System.out.println(sep1 + "\n");
    }

    // =========================================================
    //  IMPRIMIR RESUMEN EN CONSOLA
    // =========================================================

    public void imprimirResumenNave() {
        HashMap<String, Object> datos = extraerDatosCostos(this.nroVisita);
        if (datos.isEmpty()) return;

        double sumaTeus   = getDouble(datos, "teus");
        double sumaMovCont= getDouble(datos, "movCont");
        double costoEstiba= getDouble(datos, "costoEstiba");
        double costoRTG   = getDouble(datos, "costoRTG");

        System.out.println("════════ RESUMEN DE LA NAVE ════════");
        System.out.println("NOMBRE VISITA    : " + getString(datos, "nombreVisita"));
        System.out.println("VISITA           : " + getString(datos, "nroVisita"));
        System.out.println("MUELLE           : " + getString(datos, "muelle"));
        System.out.println("SEMANA           : " + getString(datos, "semana"));
        System.out.println("CUADRILLAS       : " + getDouble(datos, "cuadrillas"));
        System.out.println("LINEA            : " + getString(datos, "lineaServicio"));
        System.out.println("FECHA DE TERMINO : " + getString(datos, "fechaFin"));
        System.out.println("TIEMPO EFECTIVO  : " + getDouble(datos, "tiempoEfectivo"));
        System.out.println("MOV CONTENEDORES : " + (int) sumaMovCont);
        System.out.println("SUMA DE TEUS     : " + (int) sumaTeus);

        System.out.println("════════ RESUMEN COSTO ESTIBA ════════");
        System.out.println("COSTO ESTIBA     : " + String.format("%.2f", costoEstiba));
        System.out.println("COSTO TEU ESTIBA : " + (sumaTeus    > 0 ? String.format("%.2f", costoEstiba / sumaTeus)    : "N/A"));
        System.out.println("COSTO CONT ESTIBA: " + (sumaMovCont > 0 ? String.format("%.2f", costoEstiba / sumaMovCont) : "N/A"));

        System.out.println("════════ RESUMEN COSTO DIESEL ════════");
        System.out.println("COSTO RTG        : " + String.format("%.2f", costoRTG));
    }

    // =========================================================
    //  COPIAR CUADRILLAS A EXCEL DE NAVES
    // =========================================================

    public void copiarResumenPegarNaves() {
        LecturaExcels lecturaExcelNaves   = null;
        LecturaExcels lecturaExcelResumen = null;
        try {
            lecturaExcelNaves   = new LecturaExcels(this.rutaNavesMes);
            lecturaExcelResumen = new LecturaExcels(this.rutaExcelResumen);

            List<String> visitas     = lecturaExcelResumen.leerColumna(nombreHojaExcelResumen, colVisitaExcelResumen);
            List<String> nombresHojas= lecturaExcelNaves.obtenerNombresDeHojas();

            for (String visita : visitas) {
                if (visita == null || visita.trim().isEmpty()) continue;

                String[] partes = visita.trim().split("-");
                if (partes.length < 1) continue;
                String nroVisitaRaw = partes[0].trim();

                if (!nroVisitaRaw.matches("\\d+")) {
                    System.out.println("[SKIP] No numérico: '" + nroVisitaRaw + "'");
                    continue;
                }

                String nroVisitaSinCeros = String.valueOf(Integer.parseInt(nroVisitaRaw));
                String hojaEncontrada = null;

                for (String nombreHoja : nombresHojas) {
                    String nombreHojaSinCeros;
                    try {
                        nombreHojaSinCeros = String.valueOf(Integer.parseInt(nombreHoja.trim()));
                    } catch (NumberFormatException e) {
                        nombreHojaSinCeros = nombreHoja.trim().replaceAll("^0+", "");
                    }
                    if (nombreHojaSinCeros.equals(nroVisitaSinCeros)) {
                        hojaEncontrada = nombreHoja;
                        break;
                    }
                }

                if (hojaEncontrada == null) {
                    System.out.println("[WARN] No se encontró hoja para: " + visita);
                    continue;
                }

                List<String> columnaVisitasResumen = lecturaExcelResumen.leerColumna(
                        nombreHojaExcelResumen, colVisitaExcelResumen);

                int filaEnResumen = -1;
                for (int i = 2; i < columnaVisitasResumen.size(); i++) {
                    String v = columnaVisitasResumen.get(i).trim();
                    if (v.isEmpty() || !v.split("-")[0].trim().matches("\\d+")) continue;
                    if (String.valueOf(Integer.parseInt(v.split("-")[0].trim())).equals(nroVisitaSinCeros)) {
                        filaEnResumen = i;
                        break;
                    }
                }

                if (filaEnResumen == -1) continue;

                String valorCuadrillas = lecturaExcelResumen.leerCelda(
                        nombreHojaExcelResumen, filaEnResumen, colResumenCuadrillas);

                lecturaExcelNaves.escribirCeldaNumero(
                        hojaEncontrada, filaCuadrillas, colCuadrillas,
                        Double.parseDouble(valorCuadrillas));
                System.out.println("[OK] Cuadrillas pegadas en hoja: " + hojaEncontrada);
            }

            lecturaExcelNaves.guardar();
            System.out.println("[OK] Proceso copiarResumenPegarNaves completado.");

        } catch (Exception err) {
            System.out.println("[ERROR] copiarResumenPegarNaves: " + err.getMessage());
            err.printStackTrace();
        } finally {
            try { if (lecturaExcelNaves   != null) lecturaExcelNaves.cerrar();   } catch (Exception ignored) {}
            try { if (lecturaExcelResumen != null) lecturaExcelResumen.cerrar(); } catch (Exception ignored) {}
        }
    }

    // =========================================================
    //  UTILIDADES PRIVADAS
    // =========================================================

    private String extraerMesDesdeFecha(String fechaStr) {
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);
            LocalDateTime dt = LocalDateTime.parse(fechaStr.trim(), formatter);
            // Retorna abreviación del mes en inglés en mayúsculas: "JAN", "FEB", etc.
            return dt.getMonth().getDisplayName(
                    java.time.format.TextStyle.SHORT, Locale.ENGLISH).toUpperCase();
        } catch (Exception e) {
            System.out.println("[WARN] No se pudo extraer mes de: '" + fechaStr + "'");
            return "";
        }
    }

    private void setCeldaTexto(Row fila, int columna, String valor) {
        org.apache.poi.ss.usermodel.Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor != null ? valor : "");
    }

    private void setCeldaNumero(Row fila, int columna, double valor) {
        org.apache.poi.ss.usermodel.Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor);
    }

    private double parsearDouble(String valor) {
        if (valor == null || valor.trim().isEmpty()) return 0.0;
        try {
            return Double.parseDouble(valor.trim().replace(",", "."));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private String getString(HashMap<String, Object> datos, String clave) {
        Object val = datos.get(clave);
        return val != null ? val.toString() : "";
    }

    private double getDouble(HashMap<String, Object> datos, String clave) {
        Object val = datos.get(clave);
        if (val instanceof Double) return (Double) val;
        if (val instanceof Integer) return ((Integer) val).doubleValue();
        return 0.0;
    }

    // =========================================================
    //  SETTERS
    // =========================================================

    public void setNroVisita(String nroVisita)               { this.nroVisita = nroVisita; }
    public void setRutaNavesMes(String rutaNavesMes)          { this.rutaNavesMes = rutaNavesMes; }
    public void setRutaThroughput(String rutaThroughput)      { this.rutaThroughput = rutaThroughput; }
    public void setRutaExcelDestino(String rutaExcelDestino)  {
        this.rutaExcelDestino = rutaExcelDestino;
        this.escritorDetalle  = new EscritorDetalleNave(rutaExcelDestino);
    }
    public void setNombreHojaExcelResumen(String n)           { this.nombreHojaExcelResumen = n; }

    public void setCostoDepreciacionSTS(double costoDepreciacionSTS) {
        this.costoDepreciacionSTS = costoDepreciacionSTS;
    }

    public void setCostoDepreciacionGM(double costoDepreciacionGM) {
        this.costoDepreciacionGM = costoDepreciacionGM;
    }

    public void setCostoDepreciacionERTG(double costoDepreciacionERTG) {
        this.costoDepreciacionERTG = costoDepreciacionERTG;
    }

    public void setCostoDepreciacionRTG(double costoDepreciacionRTG) {
        this.costoDepreciacionRTG = costoDepreciacionRTG;
    }

    public void setCostoDepreciacionITV(double costoDepreciacionITV) {
        this.costoDepreciacionITV = costoDepreciacionITV;
    }

    public void setCostoDepreciacionRSK(double costoDepreciacionRSK) {
        this.costoDepreciacionRSK = costoDepreciacionRSK;
    }

    public void setCostoDepreciacionEH(double constoDepreciacionEH) {
        this.costoDepreciacionEH = constoDepreciacionEH;
    }

    public void setNombreHojaCosmos(String nombreHojaCosmos) {this.nombreHojaCosmos = nombreHojaCosmos;}

    public void setNombreHojaMoveHistory(String nombreHojaMoveHistory){this.nombreHojaMoveHistory = nombreHojaMoveHistory;}

    public void setNumMesOM(int numMesOM){this.numMesOM = numMesOM;}
}