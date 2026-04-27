package postmortem.pdf;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class PaginaIndicadores {

    // ── Colores del diseño ────────────────────────────────────
    private static final String AZUL_OSCURO  = "#1a2e4a";
    private static final String AZUL_MEDIO   = "#2e4d7b";
    private static final String AZUL_CLARO   = "#5b7fa6";
    private static final String AZUL_HEADER  = "#3a5f8a";
    private static final String VERDE_PROD   = "#c8f0c8";
    private static final String FONDO_TABLA  = "#e8f0f8";
    private static final String FONDO_PAGINA = "#f0f4f0";

    /**
     * Genera el PDF de la página de indicadores.
     *
     * @param kpi          Mapa con todos los KPIs de la visita (salida de LecturaKPIsPostMortem)
     * @param rutaSalida   Ruta completa del PDF a generar (ej: "C:/out/postmortem.pdf")
     */
    public void generarPagina(Map<String, Object> kpi, String rutaSalida) {
        String html = construirHtml(kpi);
        convertirHtmlAPdf(html, rutaSalida);
        System.out.println("[PDF] Página de indicadores generada: " + rutaSalida);
    }

    /**
     * Devuelve el HTML como String (útil para debug o preview en navegador).
     */
    public String generarHtml(Map<String, Object> kpi) {
        return construirHtml(kpi);
    }

    // ==========================================================
    // CONSTRUCCIÓN DEL HTML
    // ==========================================================
    private String construirHtml(Map<String, Object> kpi) {

        // ── Extraer datos del mapa ────────────────────────────
        String visita      = str(kpi, "Visita");
        String linea       = str(kpi, "Linea");
        String servicio    = str(kpi, "Servicio");
        String berthATA    = str(kpi, "BerthATA");
        String berthATD    = str(kpi, "BerthATD");
        String startWork   = str(kpi, "StartWork");
        String endWork     = str(kpi, "EndWork");
        double gmph        = dbl(kpi, "GMPH");
        double bmph        = dbl(kpi, "BMPH");
        double durOp       = dbl(kpi, "DuracionOp");
        double craneInt    = dbl(kpi, "CraneIntensity");
        int    totalMovs   = integer(kpi, "TotalMovimientos");

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> tablaGruas =
                (List<Map<String, Object>>) kpi.get("TablaGruas");

        // Long Gang = movs de la grúa con más movimientos
        int longGang = tablaGruas.stream()
                .mapToInt(g -> integer(g, "Movs"))
                .max().orElse(0);

        // Grúas utilizadas = cantidad de grúas STS
        int gruasUtilizadas = tablaGruas.size();

        StringBuilder sb = new StringBuilder();
        sb.append(htmlHead());
        sb.append("<body>");

        // ── Barra de título ───────────────────────────────────
        sb.append(barraTitulo());

        // ── Encabezado de nave ────────────────────────────────
        sb.append(encabezadoNave(visita, berthATA, linea, servicio,
                visita, "2", startWork, endWork));

        // ── Cuerpo principal: tabla izquierda + gráficas derecha
        sb.append("<div class='main-body'>");
        sb.append("<div class='left-panel'>");
        sb.append(tablaIndicadores(craneInt, gruasUtilizadas, durOp,
                longGang, bmph, gmph, kpi));
        sb.append("</div>"); // left-panel

        sb.append("<div class='right-panel'>");
        sb.append(graficaMovimientosEfectivos(totalMovs, tablaGruas));
        sb.append(graficaSegregacionDemoras(kpi));
        sb.append("</div>"); // right-panel

        sb.append("</div>"); // main-body
        sb.append("</body></html>");

        return sb.toString();
    }


    private String htmlHead() {
        return "<!DOCTYPE html><html><head><meta charset='UTF-8'/>"
                + "<style>" + css() + "</style></head>";

    }

    private String barraTitulo() {
        return "<div class='title-bar'>"
                + "<span class='back-arrow'>&#8592;</span>"
                + "<span class='title-text'>POST-MORTEM</span>"
                + "<span class='logo-text'>TERMINALES PORTUARIOS EUROANDINOS</span>"
                + "</div>";
    }

    private String encabezadoNave(String nombreNave, String fechaNave,
                                  String linea, String servicio,
                                  String visita, String muelle,
                                  String inicioOp, String finOp) {
        return "<div class='nav-header'>"
                + "<div class='nav-name'>"
                + "<div class='nave-titulo'>" + nombreNave + "</div>"
                + "<div class='nave-fecha'>" + fechaNave + "</div>"
                + "</div>"
                + celdaHeader(linea,   "Linea")
                + celdaHeader(servicio, "Servicio")
                + celdaHeader(visita,   "Visita")
                + celdaHeader(muelle,   "Muelle")
                + celdaHeader(inicioOp, "Inicio de Operaciones")
                + celdaHeader(finOp,    "Fin de Operaciones")
                + "</div>";
    }

    private String celdaHeader(String valor, String etiqueta) {
        return "<div class='header-cell'>"
                + "<div class='header-valor'>" + valor + "</div>"
                + "<div class='header-label'>" + etiqueta + "</div>"
                + "</div>";
    }

    // ── Tabla de indicadores (lado izquierdo) ─────────────────
    private String tablaIndicadores(double craneInt, int gruasUtilizadas,
                                    double durOp, int longGang,
                                    double bmph, double gmph,
                                    Map<String, Object> kpi) {
        StringBuilder sb = new StringBuilder();
        sb.append("<table class='kpi-table'>");

        // Encabezado Planeado / Ejecutado
        sb.append("<thead><tr>"
                + "<th class='col-indicador'></th>"
                + "<th colspan='2' class='col-ejecutado'>Ejecutado</th>"
                + "</tr></thead><tbody>");

        // Filas de indicadores
        sb.append(filaKpi("Crane Intensity", "", "QC's",
                fmt2(craneInt), "QC's", false));
        sb.append(filaKpi("Grúas utilizadas", "", "QC's",
                fmt2((double) gruasUtilizadas), "QC's", false));
        sb.append(filaKpi("Duración Operación", "", "Hrs",
                fmt2(durOp), "Hrs", false));
        sb.append(filaKpi("Long Gang", "", "Movs",
                fmt2((double) longGang), "Movs", false));
        sb.append(filaKpi("BMPH", "", "MPH",
                fmt2(bmph), "MPH", false));
        sb.append(filaKpi("GMPH", "", "MPH",
                fmt2(gmph), "MPH", false));

        // Movimientos (fondo azul oscuro)
        sb.append(filasMovimientos(kpi));

        // Demoras principales
        sb.append(filasDemoras(kpi));

        sb.append("</tbody></table>");
        return sb.toString();
    }

    private String filaKpi(String indicador, String planVal, String planUnid,
                           String ejVal, String ejUnid, boolean oscuro) {
        String cls = oscuro ? "kpi-dark" : "kpi-light";
        return "<tr class='" + cls + "'>"
                + "<td class='td-indicador'>" + indicador + "</td>"
                + "<td class='td-val-ejec'>" + ejVal + "</td>"
                + "<td class='td-unid'>" + ejUnid + "</td>"
                + "</tr>";
    }

    private String filasMovimientos(Map<String, Object> kpi) {
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> gruas =
                (List<Map<String, Object>>) kpi.get("TablaGruas");

        // Calcular descarga/embarque desde los datos de grúas si los tienes,
        // o dejar placeholder. Ajusta según tus datos reales.
        int descarga     = integer(kpi, "Descarga");
        int embarque     = integer(kpi, "Embarque");
        int restibas     = integer(kpi, "Restibas");
        int transbordos  = integer(kpi, "Transbordos");
        int totalMovs    = integer(kpi, "TotalMovimientos");

        StringBuilder sb = new StringBuilder();

        // Fila Descarga con badge FL-FM
        sb.append("<tr class='kpi-movs'>"
                + "<td class='td-movs-label'>Descarga</td>"
                + "<td class='td-movs-val'>" + descarga + "</td>"
                + "<td class='td-movs-unid'>movs</td>"
                + "<td colspan='2'><span class='badge-flor'>FL-FM</span>"
                + " <span class='badge-min'>36 min</span></td>"
                + "</tr>");

        // Fila Embarque con badge LM-LL
        sb.append("<tr class='kpi-movs'>"
                + "<td class='td-movs-label'>Embarque</td>"
                + "<td class='td-movs-val'>" + embarque + "</td>"
                + "<td class='td-movs-unid'>movs</td>"
                + "<td colspan='2'><span class='badge-lmll'>LM-LL</span>"
                + " <span class='badge-min'>44 min</span></td>"
                + "</tr>");

        sb.append("<tr class='kpi-movs'>"
                + "<td class='td-movs-label'>Restibas</td>"
                + "<td class='td-movs-val'>" + restibas + "</td>"
                + "<td class='td-movs-unid'>movs</td><td></td><td></td></tr>");

        sb.append("<tr class='kpi-movs'>"
                + "<td class='td-movs-label'>Transbordos</td>"
                + "<td class='td-movs-val'>" + transbordos + "</td>"
                + "<td class='td-movs-unid'>movs</td><td></td><td></td></tr>");

        sb.append("<tr class='kpi-total'>"
                + "<td class='td-total-label'>MOVIMIENTOS</td>"
                + "<td class='td-total-val'>" + totalMovs + "</td>"
                + "<td class='td-total-unid'>movs</td><td></td><td></td></tr>");

        return sb.toString();
    }

    private String filasDemoras(Map<String, Object> kpi) {
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> demoras =
                (List<Map<String, Object>>) kpi.get("Demoras");

        if (demoras.isEmpty()) return "";

        StringBuilder sb = new StringBuilder();

        // Encabezado de demoras
        sb.append("<tr class='demoras-header'>"
                + "<td rowspan='" + (demoras.size() + 1) + "' class='td-demoras-titulo'>"
                + "<b>Demoras<br/>Principales</b></td>"
                + "<td>Demoras</td>"
                + "<td></td>"
                + "<td>Durac (Hrs)</td>"
                + "<td>% Hrs Qc</td>"
                + "</tr>");

        for (Map<String, Object> demora : demoras) {
            sb.append("<tr class='demoras-row'>"
                    + "<td>" + str(demora, "Nombre") + "</td>"
                    + "<td></td>"
                    + "<td>" + fmt2(dbl(demora, "DurHrs")) + "</td>"
                    + "<td>" + fmt2(dbl(demora, "PctHrsQc")) + " %</td>"
                    + "</tr>");
        }

        return sb.toString();
    }

    // ── Gráfica movimientos efectivos (lado derecho) ──────────
    private String graficaMovimientosEfectivos(int totalMovs,
                                               List<Map<String, Object>> tablaGruas) {
        if (tablaGruas.isEmpty()) return "";

        int maxMovs = tablaGruas.stream()
                .mapToInt(g -> integer(g, "Movs"))
                .max().orElse(1);

        StringBuilder sb = new StringBuilder();
        sb.append("<div class='chart-section'>");
        sb.append("<div class='chart-title'>Movimientos Efectivos</div>");
        sb.append("<div class='chart-total'>" + totalMovs + "</div>");
        sb.append("<div class='bar-chart'>");

        // Una barra horizontal por grúa
        for (Map<String, Object> grua : tablaGruas) {
            String nombre = str(grua, "Grua");
            int movs      = integer(grua, "Movs");
            int pct       = (int) ((movs / (double) maxMovs) * 100);

            sb.append("<div class='bar-row'>"
                    + "<span class='bar-label'>" + nombre + "</span>"
                    + "<div class='bar-track'>"
                    + "<div class='bar-fill' style='width:" + pct + "%'>"
                    + "<span class='bar-value'>" + movs + "</span>"
                    + "</div></div></div>");
        }

        sb.append("</div></div>"); // bar-chart + chart-section
        return sb.toString();
    }

    // ── Gráfica segregación de demoras ────────────────────────
    private String graficaSegregacionDemoras(Map<String, Object> kpi) {
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> demoras =
                (List<Map<String, Object>>) kpi.get("Demoras");

        if (demoras.isEmpty()) return "";

        double totalHrs = demoras.stream()
                .mapToDouble(d -> dbl(d, "DurHrs"))
                .sum();

        StringBuilder sb = new StringBuilder();
        sb.append("<div class='chart-section'>");
        sb.append("<div class='chart-title'>Segregación de Demoras</div>");
        sb.append("<div class='chart-total'>" + fmt2(totalHrs) + "</div>");
        sb.append("<div class='demoras-chart'>");

        // Paleta de colores para las barras de demoras
        String[] colores = {"#c0392b", "#e67e22", "#f1c40f", "#27ae60", "#2980b9"};
        int colorIdx = 0;

        for (Map<String, Object> demora : demoras) {
            String nombre = str(demora, "Nombre");
            double hrs    = dbl(demora, "DurHrs");
            double pct    = dbl(demora, "PctHrsQc");
            int barPct    = totalHrs > 0 ? (int) ((hrs / totalHrs) * 100) : 0;
            String color  = colores[colorIdx % colores.length];
            colorIdx++;

            sb.append("<div class='demora-row'>"
                    + "<span class='demora-label'>" + nombre + "</span>"
                    + "<div class='demora-track'>"
                    + "<div class='demora-bar' style='width:" + barPct
                    + "%;background:" + color + "'></div>"
                    + "</div>"
                    + "<span class='demora-hrs'>" + fmt2(hrs) + " Hrs</span>"
                    + "<span class='demora-pct badge-pct' style='background:" + color + "'>"
                    + (int) Math.round(pct) + "%</span>"
                    + "</div>");
        }

        sb.append("</div></div>"); // demoras-chart + chart-section
        return sb.toString();
    }

    // ==========================================================
    // CSS
    // ==========================================================
    private String css() {
        return
                // Reset y base
                "* { box-sizing: border-box; margin: 0; padding: 0; font-family: Arial, sans-serif; font-size: 11px; }"
                        + "body { background: " + FONDO_PAGINA + "; width: 1200px; }"

                        // Barra de título
                        + ".title-bar { background: " + AZUL_OSCURO + "; color: white; "
                        + "  display: flex; align-items: center; justify-content: space-between; "
                        + "  padding: 8px 16px; font-size: 18px; font-weight: bold; }"
                        + ".back-arrow { font-size: 20px; cursor: pointer; }"
                        + ".logo-text { font-size: 13px; }"

                        // Encabezado de nave
                        + ".nav-header { display: flex; align-items: stretch; "
                        + "  border-bottom: 2px solid " + AZUL_OSCURO + "; background: white; }"
                        + ".nav-name { background: white; border: 2px solid " + AZUL_CLARO + "; "
                        + "  border-radius: 8px; padding: 6px 12px; margin: 4px; min-width: 140px; "
                        + "  display: flex; flex-direction: column; justify-content: center; }"
                        + ".nave-titulo { font-weight: bold; color: " + AZUL_MEDIO + "; font-size: 12px; }"
                        + ".nave-fecha { font-size: 10px; color: #555; }"
                        + ".header-cell { padding: 4px 12px; border-left: 1px solid #ddd; "
                        + "  display: flex; flex-direction: column; justify-content: center; }"
                        + ".header-valor { font-weight: bold; font-size: 13px; }"
                        + ".header-label { font-size: 9px; color: #777; margin-top: 2px; }"

                        // Layout principal
                        + ".main-body { display: flex; padding: 8px; gap: 12px; }"
                        + ".left-panel { flex: 0 0 480px; }"
                        + ".right-panel { flex: 1; background: #e8f5e8; border-radius: 8px; padding: 10px; }"

                        // Tabla KPI
                        + ".kpi-table { width: 100%; border-collapse: collapse; }"
                        + ".kpi-table thead tr th { background: " + AZUL_OSCURO + "; color: white; "
                        + "  padding: 6px; text-align: center; }"
                        + ".col-indicador { width: 38%; }"
                        + ".col-planeado { width: 31%; background: " + AZUL_CLARO + " !important; }"
                        + ".col-ejecutado { width: 31%; background: " + AZUL_MEDIO + " !important; }"
                        + ".kpi-light td { padding: 5px 8px; border-bottom: 1px solid #ddd; "
                        + "  background: " + FONDO_TABLA + "; }"
                        + ".kpi-dark td { padding: 5px 8px; border-bottom: 1px solid #ddd; "
                        + "  background: " + AZUL_CLARO + "; color: white; }"
                        + ".td-indicador { text-align: left; font-weight: normal; }"
                        + ".td-val-plan { text-align: right; color: #aaa; }"
                        + ".td-val-ejec { text-align: right; font-weight: bold; font-size: 12px; }"
                        + ".td-unid { text-align: left; padding-left: 4px; color: #666; }"

                        // Filas de movimientos
                        + ".kpi-movs td { padding: 4px 8px; background: " + AZUL_MEDIO + "; "
                        + "  color: white; border-bottom: 1px solid " + AZUL_OSCURO + "; }"
                        + ".td-movs-label { font-weight: normal; }"
                        + ".td-movs-val { text-align: right; font-weight: bold; }"
                        + ".td-movs-unid { text-align: left; padding-left: 4px; }"
                        + ".kpi-total td { padding: 5px 8px; background: " + AZUL_OSCURO + "; "
                        + "  color: white; font-weight: bold; font-size: 12px; }"
                        + ".td-total-val { text-align: right; }"
                        + ".td-total-unid { text-align: left; padding-left: 4px; }"

                        // Badges FL-FM / LM-LL
                        + ".badge-flor { background: " + AZUL_MEDIO + "; color: white; "
                        + "  padding: 2px 6px; border-radius: 3px; font-weight: bold; }"
                        + ".badge-lmll { background: " + AZUL_CLARO + "; color: white; "
                        + "  padding: 2px 6px; border-radius: 3px; font-weight: bold; }"
                        + ".badge-min { color: white; font-weight: bold; }"

                        // Demoras
                        + ".demoras-header td { background: #2c2c2c; color: white; "
                        + "  padding: 4px 8px; font-size: 10px; }"
                        + ".td-demoras-titulo { background: #2c2c2c; color: white; "
                        + "  text-align: center; vertical-align: middle; font-size: 11px; }"
                        + ".demoras-row td { padding: 3px 8px; background: #f5f5f5; "
                        + "  border-bottom: 1px solid #ddd; }"

                        // Panel derecho: charts
                        + ".chart-section { margin-bottom: 20px; }"
                        + ".chart-title { font-weight: bold; font-size: 13px; text-align: center; "
                        + "  margin-bottom: 4px; }"
                        + ".chart-total { text-align: center; font-size: 22px; font-weight: bold; "
                        + "  margin-bottom: 10px; }"

                        // Barras horizontales movimientos
                        + ".bar-chart { padding: 0 10px; }"
                        + ".bar-row { display: flex; align-items: center; margin-bottom: 8px; }"
                        + ".bar-label { min-width: 50px; font-weight: bold; font-size: 10px; }"
                        + ".bar-track { flex: 1; background: #ddd; border-radius: 3px; height: 24px; }"
                        + ".bar-fill { background: " + AZUL_MEDIO + "; height: 100%; border-radius: 3px; "
                        + "  display: flex; align-items: center; justify-content: flex-end; "
                        + "  padding-right: 6px; }"
                        + ".bar-value { color: white; font-weight: bold; font-size: 11px; }"

                        // Barras de demoras
                        + ".demoras-chart { padding: 0 10px; }"
                        + ".demora-row { display: flex; align-items: center; margin-bottom: 10px; gap: 8px; }"
                        + ".demora-label { min-width: 120px; font-size: 10px; text-align: right; }"
                        + ".demora-track { flex: 1; background: #eee; border-radius: 3px; height: 20px; }"
                        + ".demora-bar { height: 100%; border-radius: 3px; }"
                        + ".demora-hrs { min-width: 55px; font-size: 10px; font-weight: bold; }"
                        + ".badge-pct { color: white; padding: 2px 6px; border-radius: 3px; "
                        + "  font-weight: bold; font-size: 10px; min-width: 35px; text-align: center; }";
    }

    // ==========================================================
    // CONVERSIÓN HTML → PDF
    // ==========================================================
    private void convertirHtmlAPdf(String html, String rutaSalida) {
        try {
            // Crear carpeta si no existe
            File carpeta = new File(rutaSalida).getParentFile();
            if (carpeta != null && !carpeta.exists()) carpeta.mkdirs();

            ConverterProperties props = new ConverterProperties();

            try (OutputStream os = new FileOutputStream(rutaSalida)) {
                HtmlConverter.convertToPdf(html, os, props);
            }
        } catch (Exception e) {
            throw new RuntimeException("Error generando PDF: " + e.getMessage(), e);
        }
    }

    // ==========================================================
    // UTILIDADES
    // ==========================================================
    private String str(Map<String, Object> map, String key) {
        Object v = map.get(key);
        return v != null ? v.toString() : "";
    }

    private double dbl(Map<String, Object> map, String key) {
        Object v = map.get(key);
        if (v instanceof Number) return ((Number) v).doubleValue();
        return 0.0;
    }

    private int integer(Map<String, Object> map, String key) {
        Object v = map.get(key);
        if (v instanceof Number) return ((Number) v).intValue();
        return 0;
    }

    private String fmt2(double v) {
        return String.format(Locale.ROOT, "%.2f", v);
    }
}