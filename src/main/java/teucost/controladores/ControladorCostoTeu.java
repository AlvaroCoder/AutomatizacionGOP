package teucost.controladores;

import teucost.modelos.*;
import teucost.modelos.excels.ExcelDatosNave;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ControladorCostoTeu {

    // =========================================================
    //  TASAS DE DEPRECIACIÓN ($/hr por equipo)
    // =========================================================
    private double costoDepreciacionSTS  = 56.14;
    private double costoDepreciacionGM   = 45.54;
    private double costoDepreciacionERTG = 15.42;
    private double costoDepreciacionRTG  = 9.21;
    private double costoDepreciacionITV  = 0.65;
    private double costoDepreciacionRSK  = 2.14;
    private double costoDepreciacionEH   = 1.49;
    private double promedioVueltasITV    = 7;

    // =========================================================
    //  RUTAS DE FUENTES DE DATOS
    // =========================================================
    private String rutaThroughput;
    private String rutaControlCosmos;
    private String rutaConciliado;
    private String rutaOMMensual;
    private String rutaMoveHistory;
    private String rutaExcelDestino;
    private String rutaExcelResumen;
    private String rutaNavesMes;

    // =========================================================
    //  CONFIGURACIÓN DE HOJAS
    // =========================================================
    private String nombreHojaCosmos      = "Resumen Visitas Feb 26";
    private String nombreHojaMoveHistory = "MoveEvent";
    private String nombreHojaExcelResumen= "Resumen Visitas Feb 26";
    private int    numMesOM              = 0;

    // =========================================================
    //  CONFIGURACIÓN TABLA DESTINO — hoja 0 (0-based)
    // =========================================================
    private final int hojaDestinoCostos      = 0;
    private final int colMes                 = 0;
    private final int colNombreNave          = 1;
    private final int colVisitaTabla         = 2;
    private final int colLinea               = 3;
    private final int colFechaFin            = 4;
    private final int colSemana              = 5;
    private final int colCuadrillasTabla     = 6;
    private final int colMovCont             = 7;
    private final int colTeus                = 8;
    private final int colMuelle              = 9;
    private final int colTiempoEfectivo      = 10;
    private final int colCostoEstiba         = 11;
    private final int colCostoRTG            = 12;
    private final int colCostoReachStack     = 13;
    private final int colCostoEmptyHand      = 14;
    private final int colCostoGruaMov        = 15;
    private final int colCostoITV            = 16;
    private final int colSubtotalCosto       = 17;
    private final int colCostoSTS01          = 18;
    private final int colCostoSTS02          = 19;
    private final int colCostoSTS03          = 20;
    private final int colCostoRTG05          = 21;
    private final int colCostoRTG06          = 22;
    private final int colCostoSubtotalSTSRTG = 23;
    private final int colCostoDeprecSTS01    = 24;
    private final int colCostoDeprecSTS02    = 25;
    private final int colCostoDeprecSTS03    = 26;
    private final int colCostoDeprecGM01     = 27;
    private final int colCostoDeprecGM02     = 28;
    private final int colCostoDeprecRTG05    = 29;
    private final int colCostoDeprecRTG06    = 30;
    private final int colCostoDeprecRTG01    = 31;
    private final int colCostoDeprecRTG02    = 32;
    private final int colCostoDeprecRTG03    = 33;
    private final int colCostoDeprecRTG04    = 34;
    private final int colCostoDeprecS501     = 35;
    private final int colCostoDeprecS502     = 36;
    private final int colCostoDeprecS503     = 37;
    private final int colCostoDeprecS506     = 38;
    private final int colCostoDeprecS504     = 39;
    private final int colCostoDeprecS505     = 40;
    private final int colCostoDeprecS507     = 41;
    private final int colCostoDeprecTT       = 42;
    private final int colCostoDepreciacion   = 43;
    private final int colCostoNaveOperativo  = 44;
    private final int colCostoTeu            = 45;

    // Copiar cuadrillas
    private int    colResumenCuadrillas  = 10;
    private int    colVisitaExcelResumen = 1;
    private final int filaCuadrillas     = 24;
    private final int colCuadrillas      = 2;

    // Escritor de hoja de detalle
    private EscritorDetalleNave escritorDetalle;

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

    public void procesarVisitas(List<String> nrosVisita) {
        System.out.println("\n[INFO] ══════════════════════════════════════════════");
        System.out.println("[INFO]  INICIO — lote de " + nrosVisita.size() + " visitas");
        System.out.println("[INFO] ══════════════════════════════════════════════");

        // ── FASE 1: Cargar todas las fuentes en memoria ───────────
        System.out.println("\n[FASE 1] Cargando fuentes de datos...");

        LecturaMoveHistory moveHistory = cargarMoveHistory(nrosVisita);
        LecturaControlCosmos cosmos    = cargarCosmos(nrosVisita);
        LecturaThroughput throughput   = cargarThroughput(nrosVisita);
        LecturaConciliado conciliado   = cargarConciliado();

        LecturaOmMensual omMensual     = cargarOmMensual(cosmos, nrosVisita);

        if (omMensual == null) {
            System.out.println("[ERROR] No se pudo cargar OM Mensual. Proceso abortado.");
            return;
        }

        // Extraer tarifas del OM Mensual (son las mismas para todas las visitas del mes)
        double costoEstibaUnit  = omMensual.extraerCostoEstiba();
        HashMap<String, Double> costosDiesel    = omMensual.extraerResumenCostosDiesel();
        HashMap<String, Double> costosElectricos = omMensual.extraerResumenCostosEnergia();

        double costoPorMovRTG     = costosDiesel.getOrDefault("costoPorMovRTG",    0.0);
        double costoPorMovStacker = costosDiesel.getOrDefault("costoPorMovSTK",    0.0);
        double costoPorMovEH      = costosDiesel.getOrDefault("costoPorMovEH",     0.0);
        double costoGMGalonHora   = costosDiesel.getOrDefault("costoGMGalonHora",  0.0);
        double costoTTHora        = costosDiesel.getOrDefault("costoPorMovTT",     0.0);
        double costoUnitGalon     = costosDiesel.getOrDefault("costoUnitGalon",    0.0);
        double costoSTS01Tarifa   = costosElectricos.getOrDefault("costoEnergiaSTS01", 0.0);
        double costoSTS02Tarifa   = costosElectricos.getOrDefault("costoEnergiaSTS02", 0.0);
        double costoSTS03Tarifa   = costosElectricos.getOrDefault("costoEnergiaSTS03", 0.0);
        double costoERTG05Tarifa  = costosElectricos.getOrDefault("costoEnergiaRTG05", 0.0);
        double costoERTG06Tarifa  = costosElectricos.getOrDefault("costoEnergiaRTG06", 0.0);

        // ── FASE 2: Calcular en memoria ───────────────────────────
        System.out.println("\n[FASE 2] Calculando costos en memoria...");

        List<ResultadoNave> resultados = new ArrayList<>();
        List<String> fallidas = new ArrayList<>();

        for (String nroVisita : nrosVisita) {
            try {
                ResultadoNave resultado = calcularVisita(
                        nroVisita,
                        moveHistory, cosmos, throughput, conciliado,
                        costoEstibaUnit,
                        costoPorMovRTG, costoPorMovStacker, costoPorMovEH,
                        costoGMGalonHora, costoTTHora, costoUnitGalon,
                        costoSTS01Tarifa, costoSTS02Tarifa, costoSTS03Tarifa,
                        costoERTG05Tarifa, costoERTG06Tarifa
                );
                resultados.add(resultado);
                System.out.println("[OK] Calculada: " + resultado);

            } catch (Exception e) {
                System.out.println("[ERROR] Falló cálculo para '" + nroVisita + "': " + e.getMessage());
                e.printStackTrace();
                fallidas.add(nroVisita);
            }
        }

        // ── FASE 3: Escritura única al Excel destino ──────────────
        System.out.println("\n[FASE 3] Escribiendo " + resultados.size() + " filas al Excel destino...");

        if (!resultados.isEmpty()) {
            escribirTodasLasFilas(resultados);
            if (escritorDetalle != null) {
                // escritorDetalle.escribirTodasLasFilas(resultados);
            }
        }

        // ── Resumen final ──────────────────────────────────────────
        System.out.println("\n════════════ RESUMEN DEL LOTE ════════════");
        System.out.println(" Total visitas    : " + nrosVisita.size());
        System.out.println(" Exitosas         : " + resultados.size());
        System.out.println(" Fallidas         : " + fallidas.size());
        if (!fallidas.isEmpty()) System.out.println(" Visitas fallidas : " + fallidas);
        System.out.println("══════════════════════════════════════════");
    }

    // =========================================================
    //  FASE 1 — CARGA DE FUENTES
    // =========================================================

    private LecturaMoveHistory cargarMoveHistory(List<String> visitas) {
        try {
            LecturaMoveHistory lmh = new LecturaMoveHistory(rutaMoveHistory);
            lmh.setNombreHoja(nombreHojaMoveHistory);
            lmh.cargarDatosDeVisitas(visitas);
            System.out.println("[OK] MoveHistory cargado.");
            return lmh;
        } catch (Exception e) {
            System.out.println("[ERROR] cargarMoveHistory: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }

    private LecturaControlCosmos cargarCosmos(List<String> visitas) {
        try {
            LecturaControlCosmos lcc = new LecturaControlCosmos(rutaControlCosmos);
            lcc.setNombreHoja(nombreHojaCosmos);
            lcc.cargarDatosDeVisitas(visitas);
            System.out.println("[OK] ControlCosmos cargado.");
            return lcc;
        } catch (Exception e) {
            System.out.println("[ERROR] cargarCosmos: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }

    private LecturaThroughput cargarThroughput(List<String> visitas) {
        try {
            LecturaThroughput lt = new LecturaThroughput(rutaThroughput);
            lt.cargarDatosDeVisitas(visitas);   // debes aplicar el mismo patrón en LecturaThroughput
            System.out.println("[OK] Throughput cargado.");
            return lt;
        } catch (Exception e) {
            System.out.println("[ERROR] cargarThroughput: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }

    private LecturaConciliado cargarConciliado() {
        try {
            LecturaConciliado lc = new LecturaConciliado(rutaConciliado);
            lc.cargarTodosLosDatos();   // debes aplicar el mismo patrón en LecturaConciliado
            System.out.println("[OK] Conciliado cargado.");
            return lc;
        } catch (Exception e) {
            System.out.println("[ERROR] cargarConciliado: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }

    /**
     * Carga OM Mensual determinando el mes desde la primera visita con fecha válida,
     * o desde numMesOM si fue configurado manualmente.
     */
    private LecturaOmMensual cargarOmMensual(LecturaControlCosmos cosmos,
                                             List<String> visitas) {
        try {
            int mes = numMesOM;
            if (mes <= 0) {
                // Buscar el mes desde la primera visita con fecha válida
                for (String v : visitas) {
                    mes = cosmos.extraerNumeroMes(v);
                    if (mes > 0) break;
                }
            }
            if (mes <= 0) {
                System.out.println("[ERROR] No se pudo determinar el mes para OM Mensual.");
                return null;
            }

            LecturaOmMensual lom = new LecturaOmMensual(rutaOMMensual);
            lom.setNumMes(mes);
            lom.setNombreHojaOMDiesel("Fuel26");
            lom.setNombreHojaEnergia("WS-Energy");
            lom.setNombreHojaEstibas("Stevedoring26");
            System.out.println("[OK] OM Mensual cargado para mes: " + mes);
            return lom;
        } catch (Exception e) {
            System.out.println("[ERROR] cargarOmMensual: " + e.getMessage());
            return null;
        }
    }

    // =========================================================
    //  FASE 2 — CÁLCULO DE UNA VISITA → ResultadoNave
    // =========================================================

    private ResultadoNave calcularVisita(
            String nroVisita,
            LecturaMoveHistory moveHistory,
            LecturaControlCosmos cosmos,
            LecturaThroughput throughput,
            LecturaConciliado conciliado,
            double costoEstibaUnit,
            double costoPorMovRTG,
            double costoPorMovStacker,
            double costoPorMovEH,
            double costoGMGalonHora,
            double costoTTHora,
            double costoUnitGalon,
            double costoSTS01Tarifa,
            double costoSTS02Tarifa,
            double costoSTS03Tarifa,
            double costoERTG05Tarifa,
            double costoERTG06Tarifa) {

        System.out.println("\n[INFO] ─── Calculando visita: " + nroVisita + " ───");

        // ── Throughput ────────────────────────────────────────────
        HashMap<String, Integer> movTeus = throughput.extraerResumenMovTeus(nroVisita);
        int sumaMovCont = movTeus.getOrDefault("sumaMovContenedores", 0);
        int sumaTeus    = movTeus.getOrDefault("sumaTeus",            0);

        // ── Cosmos ────────────────────────────────────────────────
        HashMap<String, String> resumenCosmos = cosmos.extraerResumenNaveCuad(nroVisita);
        String nombreVisita      = resumenCosmos.getOrDefault("nombreVisita",      "");
        String lineaServicio     = resumenCosmos.getOrDefault("lineaServicio",     "");
        String cuadrillaStr      = resumenCosmos.getOrDefault("cuadrillas",        "0");
        String fechaFinalizacion = resumenCosmos.getOrDefault("fechaFinalizacion", "");
        String fechaInicio       = resumenCosmos.getOrDefault("fechaInicio",       "");
        String nroSemana         = resumenCosmos.getOrDefault("semana",            "");
        String tiempoEfectivoStr = resumenCosmos.getOrDefault("tiempoEfectivo",    "0");
        double cuadrillas        = parsearDouble(cuadrillaStr);
        double tiempoEfectivo    = parsearDouble(tiempoEfectivoStr);
        String mes               = cosmos.extraerMes(nroVisita);

        // ── Conciliado ────────────────────────────────────────────
        String muelle = conciliado.extraerMuelleNave(nroVisita);

        // ── MoveHistory: duraciones ───────────────────────────────
        Map<String, Double> duraciones = moveHistory.extraerResumenDuracionGruas(nroVisita);

        double durSTS01 = duraciones.getOrDefault("STS01", 0.0);
        double durSTS02 = duraciones.getOrDefault("STS02", 0.0);
        double durSTS03 = duraciones.getOrDefault("STS03", 0.0);
        double durGM01  = duraciones.getOrDefault("GM01",  0.0);
        double durGM02  = duraciones.getOrDefault("GM02",  0.0);
        double durRTG01 = duraciones.getOrDefault("RTG01", 0.0);
        double durRTG02 = duraciones.getOrDefault("RTG02", 0.0);
        double durRTG03 = duraciones.getOrDefault("RTG03", 0.0);
        double durRTG04 = duraciones.getOrDefault("RTG04", 0.0);
        double durRTG05 = duraciones.getOrDefault("RTG05", 0.0);
        double durRTG06 = duraciones.getOrDefault("RTG06", 0.0);
        double durS501  = duraciones.getOrDefault("S-501", 0.0);
        double durS502  = duraciones.getOrDefault("S-502", 0.0);
        double durS503  = duraciones.getOrDefault("S-503", 0.0);
        double durS504  = duraciones.getOrDefault("S-504", 0.0);
        double durS505  = duraciones.getOrDefault("S-505", 0.0);
        double durS506  = duraciones.getOrDefault("S-506", 0.0);
        double durS507  = duraciones.getOrDefault("S-507", 0.0);

        // ── MoveHistory: movimientos ──────────────────────────────
        int movRTG = moveHistory.extraerMovimientosRTG(nroVisita);
        int movRSK = moveHistory.extraerMovimientosStacker(nroVisita);
        int movEH  = moveHistory.extraerMovimientosEmptyHand(nroVisita);

        // Grúas activas (para cálculo ITV)
        long gruasGMActivas  = duraciones.entrySet().stream()
                .filter(e -> e.getKey().startsWith("GM")  && e.getValue() > 0.0).count();
        long gruasSTSActivas = duraciones.entrySet().stream()
                .filter(e -> e.getKey().startsWith("STS") && e.getValue() > 0.0).count();

        double itvAsignados = gruasGMActivas * 4.0 + gruasSTSActivas * 5.0;

        // ── Costos diesel ─────────────────────────────────────────
        double costoEstiba         = costoEstibaUnit * cuadrillas;
        double costoRTG            = movRTG * costoPorMovRTG;
        double costoRSK            = movRSK * costoPorMovStacker;
        double costoEH             = movEH  * costoPorMovEH;
        double costoGM             = (durGM01 + durGM02) * costoGMGalonHora;
        double galonesITVHora      = tiempoEfectivo * costoTTHora;
        double costoITV            = itvAsignados * galonesITVHora * costoUnitGalon;
        double subtotalDiesel      = costoRTG + costoRSK + costoEH + costoGM + costoITV;

        // ── Costos energía eléctrica ──────────────────────────────
        double costoTotalSTS01     = costoSTS01Tarifa  * durSTS01;
        double costoTotalSTS02     = costoSTS02Tarifa  * durSTS02;
        double costoTotalSTS03     = costoSTS03Tarifa  * durSTS03;
        double costoTotalRTG05     = costoERTG05Tarifa * durRTG05;
        double costoTotalRTG06     = costoERTG06Tarifa * durRTG06;
        double subtotalEnergia     = costoTotalSTS01 + costoTotalSTS02 + costoTotalSTS03
                + costoTotalRTG05 + costoTotalRTG06;

        // ── Depreciaciones ────────────────────────────────────────
        double depSTS01 = durSTS01 * costoDepreciacionSTS;
        double depSTS02 = durSTS02 * costoDepreciacionSTS;
        double depSTS03 = durSTS03 * costoDepreciacionSTS;
        double depGM01  = durGM01  * costoDepreciacionGM;
        double depGM02  = durGM02  * costoDepreciacionGM;
        double depRTG05 = durRTG05 * costoDepreciacionERTG;
        double depRTG06 = durRTG06 * costoDepreciacionERTG;
        double depRTG01 = durRTG01 * costoDepreciacionRTG;
        double depRTG02 = durRTG02 * costoDepreciacionRTG;
        double depRTG03 = durRTG03 * costoDepreciacionRTG;
        double depRTG04 = durRTG04 * costoDepreciacionRTG;
        double depS501  = durS501  * costoDepreciacionRSK;
        double depS502  = durS502  * costoDepreciacionRSK;
        double depS503  = durS503  * costoDepreciacionRSK;
        double depS506  = durS506  * costoDepreciacionRSK;
        double depS504  = durS504  * costoDepreciacionEH;
        double depS505  = durS505  * costoDepreciacionEH;
        double depS507  = durS507  * costoDepreciacionEH;
        double depTT    = itvAsignados * promedioVueltasITV * costoDepreciacionITV;
        double subtotalDepreciacion = depSTS01 + depSTS02 + depSTS03
                + depGM01 + depGM02
                + depRTG05 + depRTG06
                + depRTG01 + depRTG02 + depRTG03 + depRTG04
                + depS501 + depS502 + depS503 + depS506
                + depS504 + depS505 + depS507
                + depTT;

        // ── Totales ───────────────────────────────────────────────
        double costoNaveTotal = costoEstiba + subtotalDiesel + subtotalEnergia + subtotalDepreciacion;
        double costoPorTeu    = sumaTeus > 0 ? costoNaveTotal / sumaTeus : 0.0;

        // ── Construir ResultadoNave ────────────────────────────────
        return new ResultadoNave.Builder()
                .nroVisita(nroVisita)
                .nombreNave(nombreVisita)
                .mes(mes)
                .semana(nroSemana)
                .lineaServicio(lineaServicio)
                .muelle(muelle)
                .fechaInicio(fechaInicio)
                .fechaFin(fechaFinalizacion)
                .cuadrillas(cuadrillas)
                .tiempoEfectivo(tiempoEfectivo)
                .movContenedores(sumaMovCont)
                .teus(sumaTeus)
                .durSTS01(durSTS01).durSTS02(durSTS02).durSTS03(durSTS03)
                .durGM01(durGM01).durGM02(durGM02)
                .durRTG01(durRTG01).durRTG02(durRTG02)
                .durRTG03(durRTG03).durRTG04(durRTG04)
                .durRTG05(durRTG05).durRTG06(durRTG06)
                .movRTG(movRTG).movRSK(movRSK).movEH(movEH)
                .itvAsignados(itvAsignados)
                .costoUnitCuadrilla(costoEstibaUnit)
                .costoEstiba(costoEstiba)
                .costoRTG(costoRTG).costoRSK(costoRSK).costoEH(costoEH)
                .costoGM(costoGM).costoITV(costoITV)
                .subtotalDiesel(subtotalDiesel)
                .costoSTS01(costoTotalSTS01).costoSTS02(costoTotalSTS02).costoSTS03(costoTotalSTS03)
                .costoRTG05(costoTotalRTG05).costoRTG06(costoTotalRTG06)
                .subtotalEnergia(subtotalEnergia)
                .depSTS01(depSTS01).depSTS02(depSTS02).depSTS03(depSTS03)
                .depGM01(depGM01).depGM02(depGM02)
                .depRTG05(depRTG05).depRTG06(depRTG06)
                .depRTG01(depRTG01).depRTG02(depRTG02).depRTG03(depRTG03).depRTG04(depRTG04)
                .depS501(depS501).depS502(depS502).depS503(depS503).depS506(depS506)
                .depS504(depS504).depS505(depS505).depS507(depS507)
                .depTT(depTT).subtotalDepreciacion(subtotalDepreciacion)
                .costoNaveTotal(costoNaveTotal)
                .costoPorTeu(costoPorTeu)
                .build();
    }

    // =========================================================
    //  FASE 3 — ESCRITURA ÚNICA AL EXCEL DESTINO
    // =========================================================

    /**
     * Abre el Excel destino UNA sola vez, escribe todas las filas y cierra.
     */
    private void escribirTodasLasFilas(List<ResultadoNave> resultados) {
        LecturaExcels excelDestino = null;
        try {
            excelDestino = new LecturaExcels(this.rutaExcelDestino);
            for (ResultadoNave r : resultados) {
                Row fila = excelDestino.insertarFilaEnTabla(hojaDestinoCostos, 0);
                poblarFilaDestino(fila, r);
            }
            excelDestino.guardar();
            System.out.println("[OK] " + resultados.size()
                    + " filas escritas en hoja de costos.");
        } catch (Exception e) {
            System.out.println("[ERROR] escribirTodasLasFilas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (excelDestino != null) try { excelDestino.cerrar(); } catch (Exception ignored) {}
        }
    }

    /**
     * Vuelca un ResultadoNave en la fila de la tabla de costos (hoja 0).
     * Tipado directo — sin HashMap, sin casteos.
     */
    private void poblarFilaDestino(Row fila, ResultadoNave r) {
        setCeldaTexto(fila,  colMes,                 r.mes);
        setCeldaTexto(fila,  colNombreNave,           r.nombreNave);
        setCeldaTexto(fila,  colVisitaTabla,          r.nroVisita);
        setCeldaTexto(fila,  colLinea,                r.lineaServicio);
        setCeldaTexto(fila,  colFechaFin,             r.fechaFin);
        setCeldaTexto(fila,  colSemana,               r.semana);
        setCeldaNumero(fila, colCuadrillasTabla,      r.cuadrillas);
        setCeldaNumero(fila, colMovCont,              r.movContenedores);
        setCeldaNumero(fila, colTeus,                 r.teus);
        setCeldaTexto(fila,  colMuelle,               r.muelle);
        setCeldaNumero(fila, colTiempoEfectivo,       r.tiempoEfectivo);
        setCeldaNumero(fila, colCostoEstiba,          r.costoEstiba);
        setCeldaNumero(fila, colCostoRTG,             r.costoRTG);
        setCeldaNumero(fila, colCostoReachStack,      r.costoRSK);
        setCeldaNumero(fila, colCostoEmptyHand,       r.costoEH);
        setCeldaNumero(fila, colCostoGruaMov,         r.costoGM);
        setCeldaNumero(fila, colCostoITV,             r.costoITV);
        setCeldaNumero(fila, colSubtotalCosto,        r.subtotalDiesel);
        setCeldaNumero(fila, colCostoSTS01,           r.costoSTS01);
        setCeldaNumero(fila, colCostoSTS02,           r.costoSTS02);
        setCeldaNumero(fila, colCostoSTS03,           r.costoSTS03);
        setCeldaNumero(fila, colCostoRTG05,           r.costoRTG05);
        setCeldaNumero(fila, colCostoRTG06,           r.costoRTG06);
        setCeldaNumero(fila, colCostoSubtotalSTSRTG,  r.subtotalEnergia);
        setCeldaNumero(fila, colCostoDeprecSTS01,     r.depSTS01);
        setCeldaNumero(fila, colCostoDeprecSTS02,     r.depSTS02);
        setCeldaNumero(fila, colCostoDeprecSTS03,     r.depSTS03);
        setCeldaNumero(fila, colCostoDeprecGM01,      r.depGM01);
        setCeldaNumero(fila, colCostoDeprecGM02,      r.depGM02);
        setCeldaNumero(fila, colCostoDeprecRTG05,     r.depRTG05);
        setCeldaNumero(fila, colCostoDeprecRTG06,     r.depRTG06);
        setCeldaNumero(fila, colCostoDeprecRTG01,     r.depRTG01);
        setCeldaNumero(fila, colCostoDeprecRTG02,     r.depRTG02);
        setCeldaNumero(fila, colCostoDeprecRTG03,     r.depRTG03);
        setCeldaNumero(fila, colCostoDeprecRTG04,     r.depRTG04);
        setCeldaNumero(fila, colCostoDeprecS501,      r.depS501);
        setCeldaNumero(fila, colCostoDeprecS502,      r.depS502);
        setCeldaNumero(fila, colCostoDeprecS503,      r.depS503);
        setCeldaNumero(fila, colCostoDeprecS506,      r.depS506);
        setCeldaNumero(fila, colCostoDeprecS504,      r.depS504);
        setCeldaNumero(fila, colCostoDeprecS505,      r.depS505);
        setCeldaNumero(fila, colCostoDeprecS507,      r.depS507);
        setCeldaNumero(fila, colCostoDeprecTT,        r.depTT);
        setCeldaNumero(fila, colCostoDepreciacion,    r.subtotalDepreciacion);
        setCeldaNumero(fila, colCostoNaveOperativo,   r.costoNaveTotal);
        setCeldaNumero(fila, colCostoTeu,             r.costoPorTeu);
    }

    // =========================================================
    //  REPORTE EN CONSOLA — sin cambios funcionales
    // =========================================================

    public void imprimirReporteDetalladoNave(ResultadoNave r) {
        String sep1 = "═";
        String sep2 = "─";
        String sep3 = "·";

        System.out.println("\n" + sep1);
        System.out.printf("  REPORTE DETALLADO — VISITA: %-34s%n", r.nroVisita);
        System.out.println(sep1);

        System.out.println("\n  INFORMACIÓN GENERAL");
        System.out.println(sep2);
        System.out.printf("  %-22s : %s%n",       "Nombre Nave",    r.nombreNave);
        System.out.printf("  %-22s : %s%n",       "Mes",            r.mes);
        System.out.printf("  %-22s : %s%n",       "Semana",         r.semana);
        System.out.printf("  %-22s : %s%n",       "Línea Servicio", r.lineaServicio);
        System.out.printf("  %-22s : %s%n",       "Muelle",         r.muelle);
        System.out.printf("  %-22s : %s%n",       "Fecha Término",  r.fechaFin);
        System.out.printf("  %-22s : %.2f hrs%n", "Tiempo Efectivo",r.tiempoEfectivo);
        System.out.printf("  %-22s : %.0f%n",     "Cuadrillas",     r.cuadrillas);

        System.out.println("\n  MOVIMIENTOS");
        System.out.println(sep2);
        System.out.printf("  %-22s : %,.0f%n", "Mov. Contenedores", r.movContenedores);
        System.out.printf("  %-22s : %,.0f%n", "TEUs",              r.teus);

        System.out.println("\n  DURACIONES (hrs)");
        System.out.println(sep2);
        System.out.printf("  %-22s : %.2f%n", "STS01", r.durSTS01);
        System.out.printf("  %-22s : %.2f%n", "STS02", r.durSTS02);
        System.out.printf("  %-22s : %.2f%n", "STS03", r.durSTS03);
        System.out.printf("  %-22s : %.2f%n", "GM01",  r.durGM01);
        System.out.printf("  %-22s : %.2f%n", "GM02",  r.durGM02);
        System.out.printf("  %-22s : %.2f%n", "RTG01", r.durRTG01);
        System.out.printf("  %-22s : %.2f%n", "RTG02", r.durRTG02);
        System.out.printf("  %-22s : %.2f%n", "RTG03", r.durRTG03);
        System.out.printf("  %-22s : %.2f%n", "RTG04", r.durRTG04);
        System.out.printf("  %-22s : %.2f%n", "RTG05", r.durRTG05);
        System.out.printf("  %-22s : %.2f%n", "RTG06", r.durRTG06);

        System.out.println("\n  COSTO ESTIBA");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "Total Estiba",  r.costoEstiba);
        System.out.printf("  %-22s : $ %,.2f%n", "Costo / TEU",
                r.teus > 0 ? r.costoEstiba / r.teus : 0);

        System.out.println("\n  COSTO DIESEL");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG",            r.costoRTG);
        System.out.printf("  %-22s : $ %,.2f%n", "Reach Stacker",  r.costoRSK);
        System.out.printf("  %-22s : $ %,.2f%n", "Empty Handler",  r.costoEH);
        System.out.printf("  %-22s : $ %,.2f%n", "Grúa Móvil",     r.costoGM);
        System.out.printf("  %-22s : $ %,.2f%n", "ITV",            r.costoITV);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL DIESEL",r.subtotalDiesel);

        System.out.println("\n  ENERGÍA ELÉCTRICA");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "STS01",          r.costoSTS01);
        System.out.printf("  %-22s : $ %,.2f%n", "STS02",          r.costoSTS02);
        System.out.printf("  %-22s : $ %,.2f%n", "STS03",          r.costoSTS03);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG05",         r.costoRTG05);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG06",         r.costoRTG06);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL ENERGÍA",r.subtotalEnergia);

        System.out.println("\n  DEPRECIACIONES");
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "STS01",   r.depSTS01);
        System.out.printf("  %-22s : $ %,.2f%n", "STS02",   r.depSTS02);
        System.out.printf("  %-22s : $ %,.2f%n", "STS03",   r.depSTS03);
        System.out.printf("  %-22s : $ %,.2f%n", "GM01",    r.depGM01);
        System.out.printf("  %-22s : $ %,.2f%n", "GM02",    r.depGM02);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG05",  r.depRTG05);
        System.out.printf("  %-22s : $ %,.2f%n", "eRTG06",  r.depRTG06);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG01",   r.depRTG01);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG02",   r.depRTG02);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG03",   r.depRTG03);
        System.out.printf("  %-22s : $ %,.2f%n", "RTG04",   r.depRTG04);
        System.out.printf("  %-22s : $ %,.2f%n", "S-501",   r.depS501);
        System.out.printf("  %-22s : $ %,.2f%n", "S-502",   r.depS502);
        System.out.printf("  %-22s : $ %,.2f%n", "S-503",   r.depS503);
        System.out.printf("  %-22s : $ %,.2f%n", "S-506",   r.depS506);
        System.out.printf("  %-22s : $ %,.2f%n", "S-504",   r.depS504);
        System.out.printf("  %-22s : $ %,.2f%n", "S-505",   r.depS505);
        System.out.printf("  %-22s : $ %,.2f%n", "S-507",   r.depS507);
        System.out.printf("  %-22s : $ %,.2f%n", "TT (ITV)",r.depTT);
        System.out.println(sep3);
        System.out.printf("  %-22s : $ %,.2f%n", "SUBTOTAL DEPREC.", r.subtotalDepreciacion);

        System.out.println("\n" + sep1);
        System.out.println("  RESUMEN FINAL");
        System.out.println(sep1);
        System.out.printf("  %-22s : $ %,.2f%n", "Estiba",            r.costoEstiba);
        System.out.printf("  %-22s : $ %,.2f%n", "Diesel",            r.subtotalDiesel);
        System.out.printf("  %-22s : $ %,.2f%n", "Energía Eléctrica", r.subtotalEnergia);
        System.out.printf("  %-22s : $ %,.2f%n", "Depreciación",      r.subtotalDepreciacion);
        System.out.println(sep2);
        System.out.printf("  %-22s : $ %,.2f%n", "COSTO NAVE TOTAL",  r.costoNaveTotal);
        System.out.println(sep1);
        System.out.printf("  %-22s : $ %,.2f%n", "COSTO POR TEU",     r.costoPorTeu);
        System.out.println(sep1 + "\n");
    }

    // =========================================================
    //  COPIAR CUADRILLAS — sin cambios
    // =========================================================

    public void copiarResumenPegarNaves() {
        LecturaExcels lecturaExcelNaves   = null;
        LecturaExcels lecturaExcelResumen = null;
        try {
            lecturaExcelNaves   = new LecturaExcels(this.rutaNavesMes);
            lecturaExcelResumen = new LecturaExcels(this.rutaExcelResumen);

            List<String> visitas      = lecturaExcelResumen.leerColumna(nombreHojaExcelResumen, colVisitaExcelResumen);
            List<String> nombresHojas = lecturaExcelNaves.obtenerNombresDeHojas();

            for (String visita : visitas) {
                if (visita == null || visita.trim().isEmpty()) continue;
                String[] partes = visita.trim().split("-");
                if (partes.length < 1) continue;
                String nroVisitaRaw = partes[0].trim();
                if (!nroVisitaRaw.matches("\\d+")) continue;

                String nroSinCeros = String.valueOf(Integer.parseInt(nroVisitaRaw));
                String hojaEncontrada = null;

                for (String nombreHoja : nombresHojas) {
                    String hojaSinCeros;
                    try { hojaSinCeros = String.valueOf(Integer.parseInt(nombreHoja.trim())); }
                    catch (NumberFormatException e) { hojaSinCeros = nombreHoja.trim().replaceAll("^0+", ""); }
                    if (hojaSinCeros.equals(nroSinCeros)) { hojaEncontrada = nombreHoja; break; }
                }
                if (hojaEncontrada == null) { System.out.println("[WARN] Sin hoja para: " + visita); continue; }

                List<String> colVisitas = lecturaExcelResumen.leerColumna(nombreHojaExcelResumen, colVisitaExcelResumen);
                int filaEnResumen = -1;
                for (int i = 2; i < colVisitas.size(); i++) {
                    String v = colVisitas.get(i).trim();
                    if (v.isEmpty() || !v.split("-")[0].trim().matches("\\d+")) continue;
                    if (String.valueOf(Integer.parseInt(v.split("-")[0].trim())).equals(nroSinCeros)) {
                        filaEnResumen = i; break;
                    }
                }
                if (filaEnResumen == -1) continue;

                String valorCuadrillas = lecturaExcelResumen.leerCelda(
                        nombreHojaExcelResumen, filaEnResumen, colResumenCuadrillas);
                lecturaExcelNaves.escribirCeldaNumero(
                        hojaEncontrada, filaCuadrillas, colCuadrillas,
                        Double.parseDouble(valorCuadrillas));
                System.out.println("[OK] Cuadrillas → hoja: " + hojaEncontrada);
            }
            lecturaExcelNaves.guardar();

        } catch (Exception e) {
            System.out.println("[ERROR] copiarResumenPegarNaves: " + e.getMessage());
        } finally {
            try { if (lecturaExcelNaves   != null) lecturaExcelNaves.cerrar();   } catch (Exception ignored) {}
            try { if (lecturaExcelResumen != null) lecturaExcelResumen.cerrar(); } catch (Exception ignored) {}
        }
    }

    // =========================================================
    //  UTILIDADES PRIVADAS
    // =========================================================

    private double parsearDouble(String valor) {
        if (valor == null || valor.trim().isEmpty()) return 0.0;
        try { return Double.parseDouble(valor.trim().replace(",", ".")); }
        catch (NumberFormatException e) { return 0.0; }
    }

    private void setCeldaTexto(Row fila, int col, String valor) {
        org.apache.poi.ss.usermodel.Cell c = fila.getCell(col);
        if (c == null) c = fila.createCell(col);
        c.setCellValue(valor != null ? valor : "");
    }

    private void setCeldaNumero(Row fila, int col, double valor) {
        org.apache.poi.ss.usermodel.Cell c = fila.getCell(col);
        if (c == null) c = fila.createCell(col);
        c.setCellValue(valor);
    }

    // =========================================================
    //  SETTERS
    // =========================================================

    public void setRutaExcelDestino(String ruta) {
        this.rutaExcelDestino = ruta;
        this.escritorDetalle  = new EscritorDetalleNave(ruta);
    }
    public void setRutaThroughput(String r)         { this.rutaThroughput    = r; }
    public void setRutaNavesMes(String r)            { this.rutaNavesMes      = r; }
    public void setNombreHojaCosmos(String n)        { this.nombreHojaCosmos  = n; }
    public void setNombreHojaMoveHistory(String n)   { this.nombreHojaMoveHistory = n; }
    public void setNombreHojaExcelResumen(String n)  { this.nombreHojaExcelResumen = n; }
    public void setNumMesOM(int n)                   { this.numMesOM          = n; }
    public void setCostoDepreciacionSTS(double v)    { this.costoDepreciacionSTS  = v; }
    public void setCostoDepreciacionGM(double v)     { this.costoDepreciacionGM   = v; }
    public void setCostoDepreciacionERTG(double v)   { this.costoDepreciacionERTG = v; }
    public void setCostoDepreciacionRTG(double v)    { this.costoDepreciacionRTG  = v; }
    public void setCostoDepreciacionITV(double v)    { this.costoDepreciacionITV  = v; }
    public void setCostoDepreciacionRSK(double v)    { this.costoDepreciacionRSK  = v; }
    public void setCostoDepreciacionEH(double v)     { this.costoDepreciacionEH   = v; }
}