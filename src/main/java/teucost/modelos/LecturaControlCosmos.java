package teucost.modelos;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.time.temporal.WeekFields;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

public class LecturaControlCosmos {
    private String rutaControlCosmos;
    private String nroVisita;
    private String nombreHojaCosmos = "Resumen Visitas Feb 26";

    private final int colNroVisita = 1;
    private final int colNombreVisita = 0;
    private final int colLineaVisita = 2;
    private final int colServicioVisita = 3;
    private final int colCuadrillaVisita = 8;
    private final int colFinTrabajo = 5;
    private final int colInicioTrabajo = 4;

    public LecturaControlCosmos(String rutaControlCosmos) {
        this.rutaControlCosmos = rutaControlCosmos;
    }

    // ✅ Bug 2 corregido: retorna HashMap<String, String> en lugar de HashMap<String, Integer>
    public HashMap<String, String> extraerResumenNaveCuad() {
        // ✅ Bug 1 corregido: inicializar el mapa, no dejarlo null
        HashMap<String, String> jsonResumenNave = new HashMap<>();
        LecturaExcels excelCosmos = null;

        try {
            excelCosmos = new LecturaExcels(this.rutaControlCosmos);
            List<String> columnaVisitas = excelCosmos.leerColumna(nombreHojaCosmos, colNroVisita);

            for (int i = 2; i < columnaVisitas.size(); i++) {
                String visitaNave = columnaVisitas.get(i).trim();
                if (visitaNave.isEmpty()) continue;

                if (visitaNave.equals(nroVisita)) {
                    String nombreVisita  = excelCosmos.leerCelda(nombreHojaCosmos, i, colNombreVisita);
                    String lineaServicio = excelCosmos.leerCelda(nombreHojaCosmos, i, colLineaVisita)
                            + "-" + excelCosmos.leerCelda(nombreHojaCosmos, i, colServicioVisita);
                    String cuadrillas = excelCosmos.leerCelda(nombreHojaCosmos, i, colCuadrillaVisita);
                    String fechaFin = excelCosmos.leerCelda(nombreHojaCosmos, i, colFinTrabajo);
                    String fechaInicio = excelCosmos.leerCelda(nombreHojaCosmos, i, colInicioTrabajo);

                    String nroSemana = "";
                    String tiempoEfectivo = "";

                    try {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy HH:mm", Locale.ENGLISH);

                        LocalDateTime dtInicio = LocalDateTime.parse(fechaInicio.trim(), formatter);
                        LocalDateTime dtFin    = LocalDateTime.parse(fechaFin.trim(), formatter);

                        // ✅ Nro de semana del año según ISO 8601
                        int semana = dtFin.toLocalDate().get(WeekFields.ISO.weekOfWeekBasedYear());
                        nroSemana  = String.valueOf(semana);

                        // ✅ Tiempo efectivo en horas con decimales (equivalente a (fechaFin - fechaInicio) * 24 de Excel)
                        long minutos = ChronoUnit.MINUTES.between(dtInicio, dtFin);
                        double horas = minutos / 60.0;
                        tiempoEfectivo = String.format("%.2f", horas); // ej: "18.75"

                    } catch (Exception e) {
                        System.out.println("[WARN] No se pudo parsear fechas → " + e.getMessage());
                    }

                    jsonResumenNave.put("semana", nroSemana);
                    jsonResumenNave.put("nombreVisita", nombreVisita);
                    jsonResumenNave.put("lineaServicio", lineaServicio);
                    jsonResumenNave.put("cuadrillas", cuadrillas);
                    jsonResumenNave.put("fechaFinalizacion", fechaFin);
                    jsonResumenNave.put("fechaInicio", fechaInicio);
                    jsonResumenNave.put("tiempoEfectivo", tiempoEfectivo);
                    break;
                }
            }

        } catch (Exception err) {
            System.out.println("[ERROR] extraerResumenNaveCuad: " + err.getMessage());
            err.printStackTrace();
        } finally {
            if (excelCosmos != null) {
                try { excelCosmos.cerrar(); } catch (Exception ignored) {}
            }
        }

        return jsonResumenNave;
    }

    public void setNroVisita(String nroVisita) { this.nroVisita = nroVisita; }
    public void setNombreHojaCosmos(String nombreHojaCosmos) { this.nombreHojaCosmos = nombreHojaCosmos; }
    public void setRutaControlCosmos(String rutaControlCosmos) { this.rutaControlCosmos = rutaControlCosmos; }
}