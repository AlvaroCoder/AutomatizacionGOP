package postmortem.pdf;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class PaginaIndicadores {

    // ── CSS embebido (tu styleTabla.css completo) ─────────────
    private static final String CSS =
            "body { font-family: Arial, sans-serif; font-size: 12px; margin: 0; padding: 0; }"
                    + ".pm-page { margin: 0; background: #fff; width: 100%; }"

                    + ".title-bar { background: #162947; color: #fff; display: flex; align-items: center;"
                    + "  justify-content: space-between; padding: 8px 16px; }"
                    + ".title-bar .arrow { font-size: 16px; color: #7eb8f7; }"
                    + ".title-bar .title { font-size: 15px; font-weight: bold; letter-spacing: 3px;"
                    + "  text-transform: uppercase; }"
                    + ".title-bar .logo { font-size: 10px; color: #7eb8f7; text-align: right; line-height: 1.5; }"

                    + ".nav-datos-nave { display: flex; flex-direction: row;"
                    + "  justify-content: space-evenly; flex: 1; }"
                    + ".nav-header { display: flex; align-items: stretch;"
                    + "  border-bottom: 1.5px solid #162947; background: #fff; }"
                    + ".nav-name { border: 1.5px solid #5b7fa6; border-radius: 8px; padding: 6px 12px;"
                    + "  margin: 6px; min-width: 140px; display: flex; flex-direction: column;"
                    + "  text-align: center; }"
                    + ".nav-name .nave { font-weight: bold; color: #2e4d7b; font-size: 12px; }"
                    + ".nav-name .fecha { font-size: 10px; color: #666; }"
                    + ".hcell { padding: 5px 14px; border-left: 0.5px solid #ddd;"
                    + "  display: flex; flex-direction: column; justify-content: center; }"
                    + ".hcell .hval { font-weight: bold; font-size: 12px; color: #111; }"
                    + ".hcell .hlbl { font-size: 9px; color: #888; margin-top: 2px;"
                    + "  text-transform: uppercase; letter-spacing: 0.5px; }"

                    + ".main-body { display: flex; }"
                    + ".left-col { width: 46%; border-right: 0.5px solid #ddd; }"
                    + ".right-col { flex: 1; background: #eef6ee; padding: 20px 16px; }"

                    + ".kpi-table { width: 100%; border-collapse: collapse; }"
                    + ".kpi-table thead tr th { background: #162947; color: #fff; padding: 8px 12px;"
                    + "  text-align: center; font-size: 11px; font-weight: bold; letter-spacing: 2px;"
                    + "  text-transform: uppercase; }"

                    + ".sec-divider td { background: #2a4472; color: #a8c4f0; font-size: 9px;"
                    + "  font-weight: bold; letter-spacing: 2px; text-transform: uppercase;"
                    + "  padding: 3px 12px; }"

                    + ".krow { border-bottom: 0.5px solid #e0e7f0; }"
                    + ".krow:nth-child(odd) td { background: #eef3fb; }"
                    + ".krow:nth-child(even) td { background: #fff; }"

                    + ".col-lbl { padding: 7px 12px; text-align: left; color: #2d3a52; width: 55%; }"
                    + ".col-val { padding: 7px 6px; text-align: right; font-weight: bold;"
                    + "  font-size: 12px; color: #162947; width: 22%; }"
                    + ".col-uni { padding: 7px 12px 7px 4px; font-size: 10px; color: #6b82a8; width: 23%; }"

                    + ".krow.gkpi  .col-lbl { border-left: 2.5px solid #3a6fd8; }"
                    + ".krow.gmov  .col-lbl { border-left: 2.5px solid #1e8c5a; }"
                    + ".krow.gtime .col-lbl { border-left: 2.5px solid #c47a1e; }"

                    + ".row-total td { background: #162947 !important; color: #fff;"
                    + "  font-weight: bold; padding: 8px 12px; }"
                    + ".row-total .col-val { color: #7eb8f7; text-align: right; }"
                    + ".row-total .col-uni { color: #7eb8f7; }"

                    + ".dem-row td { background: #2a2a2a; color: #ddd; padding: 5px 12px;"
                    + "  border-bottom: 0.5px solid #333; font-size: 11px; }"
                    + ".dem-val { text-align: right; }"
                    + ".dem-pct { color: #aaa; }"

                    + ".chart-block { margin-bottom: 28px; }"
                    + ".chart-title { font-weight: bold; font-size: 13px; text-align: center;"
                    + "  margin-bottom: 4px; }"
                    + ".chart-total { text-align: center; font-size: 28px; font-weight: bold;"
                    + "  margin-bottom: 14px; }"
                    + ".chart-total-sm { font-size: 20px; }"

                    + ".bar-row { display: flex; align-items: center; margin-bottom: 8px; gap: 8px; }"
                    + ".bar-lbl { min-width: 48px; font-size: 10px; font-weight: bold;"
                    + "  color: #555; text-align: right; }"
                    + ".bar-track { flex: 1; background: #d0dde8; height: 22px; }"
                    + ".bar-fill { height: 100%; display: flex; align-items: center;"
                    + "  justify-content: flex-end; padding-right: 7px; }"
                    + ".bar-num { color: #fff; font-weight: bold; font-size: 11px; }"

                    + ".dem-bar-row { display: flex; align-items: center; margin-bottom: 10px; gap: 8px; }"
                    + ".dem-bar-lbl { min-width: 110px; font-size: 10px; color: #555; text-align: right; }"
                    + ".dem-bar-track { flex: 1; background: #c8d8c8; height: 18px; }"
                    + ".dem-bar-fill { height: 100%; }"
                    + ".dem-hrs { min-width: 52px; font-size: 10px; font-weight: bold; }"
                    + ".dem-badge { color: #fff; padding: 2px 7px; font-size: 10px; font-weight: bold;"
                    + "  min-width: 34px; text-align: center; border-radius: 3px; }"

                    + ".axis-labels { display: flex; justify-content: space-between; font-size: 9px;"
                    + "  color: #888; margin-top: 4px; }"
                    + ".axis-title { text-align: center; font-size: 9px; color: #888; margin-top: 4px; }";

    // ==========================================================
    // PUNTO DE ENTRADA
    // ==========================================================
    public void generarPagina(Map<String, Object> kpi, String rutaSalida) {
        String html = construirHtml(kpi);
        convertirHtmlAPdf(html, rutaSalida);
        System.out.println("[PDF] Página indicadores generada: " + rutaSalida);
    }

    /** Útil para previsualizar en navegador antes de convertir a PDF */
    public void guardarHtml(Map<String, Object> kpi, String rutaHtml) {
        try {
            String html = construirHtml(kpi);
            try (Writer w = new OutputStreamWriter(
                    new FileOutputStream(rutaHtml), StandardCharsets.UTF_8)) {
                w.write(html);
            }
            System.out.println("[HTML] Guardado en: " + rutaHtml);
        } catch (Exception e) {
            throw new RuntimeException("Error guardando HTML: " + e.getMessage(), e);
        }
    }

    // ==========================================================
    // CONSTRUCCIÓN DEL HTML
    // ==========================================================
    private String construirHtml(Map<String, Object> kpi) {

        // ── Datos del encabezado ──────────────────────────────
        String nombreNave = str(kpi, "NombreNave");
        String fechaNave  = str(kpi, "FechaNave");
        String linea      = str(kpi, "Linea");
        String servicio   = str(kpi, "Servicio");
        String visita     = str(kpi, "Visita");
        String muelle     = str(kpi, "Muelle");
        String inicioOp   = str(kpi, "StartWork");
        String finOp      = str(kpi, "EndWork");

        // ── KPIs ─────────────────────────────────────────────
        String craneInt   = fmt(dbl(kpi, "CraneIntensity"));
        String gruas      = fmt((double) integer(kpi, "GruasUtilizadas"));
        String durOp      = fmt(dbl(kpi, "DuracionOp"));
        String longGang   = fmt((double) integer(kpi, "LongGang"));
        String gmph       = fmt(dbl(kpi, "GMPH"));
        String bmph       = fmt(dbl(kpi, "BMPH"));

        // ── Movimientos ───────────────────────────────────────
        int descarga    = integer(kpi, "Descarga");
        int embarque    = integer(kpi, "Embarque");
        int restibas    = integer(kpi, "Restibas");
        int transbordos = integer(kpi, "Transbordos");
        int totalMovs   = integer(kpi, "TotalMovimientos");

        // ── Tiempos ───────────────────────────────────────────
        int flfm = integer(kpi, "FL_FM_Min");
        int lmll = integer(kpi, "LM_LL_Min");

        // ── Grúas y demoras ───────────────────────────────────
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> tablaGruas =
                (List<Map<String, Object>>) kpi.get("TablaGruas");

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> demoras =
                (List<Map<String, Object>>) kpi.get("Demoras");

        double totalHrsDemoras = demoras.stream()
                .mapToDouble(d -> dbl(d, "DurHrs")).sum();

        // ── Armar HTML ────────────────────────────────────────
        return "<!DOCTYPE html><html lang='es'><head>"
                + "<meta charset='UTF-8'/>"
                + "<style>" + CSS + "</style>"
                + "</head><body>"
                + "<div class='pm-page'>"
                + barraTitulo()
                + encabezadoNave(nombreNave, fechaNave, linea, servicio,
                visita, muelle, inicioOp, finOp)
                + "<div class='main-body'>"
                + "<div class='left-col'>"
                + tablaKpi(craneInt, gruas, durOp, longGang, gmph, bmph,
                descarga, embarque, restibas, transbordos, totalMovs,
                flfm, lmll, demoras)
                + "</div>"
                + "<div class='right-col'>"
                + graficaMovimientos(totalMovs, tablaGruas)
                + graficaDemoras(totalHrsDemoras, demoras)
                + "</div>"
                + "</div>"
                + "</div>"
                + "</body></html>";
    }

    // ==========================================================
    // SECCIONES HTML
    // ==========================================================

    private String barraTitulo() {
        return "<div class='title-bar'>"
                + "<span class='arrow'>&#8592;</span>"
                + "<span class='title'>Post-Mortem</span>"
                + "<span class='logo'>&#8776; EUROANDINOS<br/>Terminales Portuarios</span>"
                + "</div>";
    }

    private String encabezadoNave(String nombreNave, String fechaNave,
                                  String linea,     String servicio,
                                  String visita,    String muelle,
                                  String inicioOp,  String finOp) {
        return "<div class='nav-header'>"
                + "<div class='nav-name'>"
                + "<span class='nave'>" + nombreNave + "</span>"
                + "<span class='fecha'>" + fechaNave  + "</span>"
                + "</div>"
                + "<div class='nav-datos-nave'>"
                + hcell(linea,    "Linea")
                + hcell(servicio, "Servicio")
                + hcell(visita,   "Visita")
                + hcell(muelle,   "Muelle")
                + hcell(inicioOp, "Inicio de operaciones")
                + hcell(finOp,    "Fin de operaciones")
                + "</div>"
                + "</div>";
    }

    private String hcell(String valor, String etiqueta) {
        return "<div class='hcell'>"
                + "<div class='hval'>" + valor    + "</div>"
                + "<div class='hlbl'>" + etiqueta + "</div>"
                + "</div>";
    }

    private String tablaKpi(String craneInt, String gruas, String durOp,
                            String longGang, String gmph,  String bmph,
                            int descarga,   int embarque,  int restibas,
                            int transbordos, int totalMovs,
                            int flfm,        int lmll,
                            List<Map<String, Object>> demoras) {
        StringBuilder sb = new StringBuilder();
        sb.append("<table class='kpi-table'><thead><tr><th colspan='3'>Ejecutado</th></tr></thead><tbody>");

        // Grupo: indicadores de grúa
        sb.append(secDivider("Indicadores de gr&#250;a"));
        sb.append(krow("gkpi", "Crane Intensity",   craneInt, "QC's"));
        sb.append(krow("gkpi", "Gr&#250;as utilizadas", gruas,  "QC's"));
        sb.append(krow("gkpi", "Duraci&#243;n operaci&#243;n", durOp, "Hrs"));
        sb.append(krow("gkpi", "Long Gang",          longGang, "Movs"));
        sb.append(krow("gkpi", "GMPH",               gmph,     "MPH"));
        sb.append(krow("gkpi", "BMPH",               bmph,     "MPH"));

        // Grupo: movimientos
        sb.append(secDivider("Movimientos"));
        sb.append(krow("gmov", "Descarga",    String.valueOf(descarga),    "Movs"));
        sb.append(krow("gmov", "Embarque",    String.valueOf(embarque),    "Movs"));
        sb.append(krow("gmov", "Reestibas",   String.valueOf(restibas),    "Movs"));
        sb.append(krow("gmov", "Transbordos", String.valueOf(transbordos), "Movs"));
        sb.append("<tr class='row-total'>"
                + "<td class='col-lbl'>Movimientos</td>"
                + "<td class='col-val'>" + totalMovs + "</td>"
                + "<td class='col-uni'>Movs</td></tr>");

        // Grupo: tiempos
        sb.append(secDivider("Tiempos"));
        sb.append(krow("gtime", "First Line to First Move", String.valueOf(flfm), "Min"));
        sb.append(krow("gtime", "Last Move to Last Line",   String.valueOf(lmll), "Min"));

        // Grupo: demoras
        sb.append(secDivider("Demoras principales"));
        for (Map<String, Object> d : demoras) {
            String nombre = str(d, "Nombre");
            String hrs    = fmt(dbl(d, "DurHrs"));
            String pct    = fmt(dbl(d, "PctHrsQc"));
            sb.append("<tr class='dem-row'>"
                    + "<td>" + nombre + "</td>"
                    + "<td class='dem-val'>" + hrs + "</td>"
                    + "<td class='dem-pct'>Hrs &nbsp; " + pct + "%</td>"
                    + "</tr>");
        }

        sb.append("</tbody></table>");
        return sb.toString();
    }

    private String secDivider(String label) {
        return "<tr class='sec-divider'><td colspan='3'>" + label + "</td></tr>";
    }

    private String krow(String grupo, String label, String valor, String unidad) {
        return "<tr class='krow " + grupo + "'>"
                + "<td class='col-lbl'>" + label  + "</td>"
                + "<td class='col-val'>" + valor  + "</td>"
                + "<td class='col-uni'>" + unidad + "</td>"
                + "</tr>";
    }

    private String graficaMovimientos(int totalMovs,
                                      List<Map<String, Object>> tablaGruas) {
        if (tablaGruas.isEmpty()) return "";

        int maxMovs = tablaGruas.stream()
                .mapToInt(g -> integer(g, "Movs")).max().orElse(1);

        StringBuilder sb = new StringBuilder();
        sb.append("<div class='chart-block'>");
        sb.append("<div class='chart-title'>Movimientos efectivos</div>");
        sb.append("<div class='chart-total'>" + totalMovs + "</div>");

        // Paleta azul para las grúas
        String[] colores = {"#2e4d7b", "#5b7fa6", "#3a6fd8", "#162947"};
        int ci = 0;

        for (Map<String, Object> grua : tablaGruas) {
            String nombre = str(grua, "Grua");
            int movs      = integer(grua, "Movs");
            int pct       = (int) ((movs / (double) maxMovs) * 100);
            String color  = colores[ci++ % colores.length];

            sb.append("<div class='bar-row'>"
                    + "<span class='bar-lbl'>" + nombre + "</span>"
                    + "<div class='bar-track'>"
                    + "<div class='bar-fill' style='width:" + pct
                    + "%;background:" + color + "'>"
                    + "<span class='bar-num'>" + movs + "</span>"
                    + "</div></div></div>");
        }

        // Eje X dinámico
        sb.append("<div class='axis-labels' style='padding-left:56px'>");
        int step = Math.max(1, (maxMovs / 4));
        for (int i = 0; i <= 4; i++) sb.append("<span>" + (step * i) + "</span>");
        sb.append("</div>");

        sb.append("</div>");
        return sb.toString();
    }

    private String graficaDemoras(double totalHrs,
                                  List<Map<String, Object>> demoras) {
        if (demoras.isEmpty()) return "";

        String[] colores = {"#c0392b", "#e67e22", "#f1c40f", "#27ae60", "#2980b9"};
        int ci = 0;

        StringBuilder sb = new StringBuilder();
        sb.append("<div class='chart-block'>");
        sb.append("<div class='chart-title'>Segregaci&#243;n de demoras</div>");
        sb.append("<div class='chart-total chart-total-sm'>"
                + fmt(totalHrs) + " Hrs</div>");

        for (Map<String, Object> d : demoras) {
            String nombre = str(d, "Nombre");
            double hrs    = dbl(d, "DurHrs");
            double pct    = dbl(d, "PctHrsQc");
            int barPct    = totalHrs > 0 ? (int) ((hrs / totalHrs) * 100) : 0;
            String color  = colores[ci++ % colores.length];

            sb.append("<div class='dem-bar-row'>"
                    + "<span class='dem-bar-lbl'>" + nombre + "</span>"
                    + "<div class='dem-bar-track'>"
                    + "<div class='dem-bar-fill' style='width:" + barPct
                    + "%;background:" + color + "'></div></div>"
                    + "<span class='dem-hrs'>" + fmt(hrs) + " Hrs</span>"
                    + "<span class='dem-badge' style='background:" + color + "'>"
                    + (int) Math.round(pct) + "%</span>"
                    + "</div>");
        }

        // Eje X fijo (en horas)
        sb.append("<div class='axis-labels' style='padding-left:118px'>"
                + "<span>0,0</span><span>0,1</span><span>0,2</span>"
                + "<span>0,3</span><span>0,4</span></div>");
        sb.append("<div class='axis-title'>Horas</div>");
        sb.append("</div>");
        return sb.toString();
    }

    // ==========================================================
    // CONVERSIÓN HTML → PDF  (hoja horizontal A4)
    // ==========================================================
    private void convertirHtmlAPdf(String html, String rutaSalida) {
        try {
            File carpeta = new File(rutaSalida).getParentFile();
            if (carpeta != null && !carpeta.exists()) carpeta.mkdirs();

            // PageSize.A4.rotate() = A4 apaisado (842 x 595 pts)
            PdfWriter   writer  = new PdfWriter(rutaSalida);
            PdfDocument pdfDoc  = new PdfDocument(writer);
            pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

            ConverterProperties props = new ConverterProperties();

            try (InputStream is = new ByteArrayInputStream(
                    html.getBytes(StandardCharsets.UTF_8))) {
                HtmlConverter.convertToPdf(is, pdfDoc, props);
            }

        } catch (Exception e) {
            throw new RuntimeException("Error generando PDF: " + e.getMessage(), e);
        }
    }

    // ==========================================================
    // UTILIDADES
    // ==========================================================
    private String str(Map<String, Object> m, String k) {
        Object v = m.get(k); return v != null ? v.toString() : "";
    }
    private double dbl(Map<String, Object> m, String k) {
        Object v = m.get(k);
        return v instanceof Number ? ((Number) v).doubleValue() : 0.0;
    }
    private int integer(Map<String, Object> m, String k) {
        Object v = m.get(k);
        return v instanceof Number ? ((Number) v).intValue() : 0;
    }
    private String fmt(double v) {
        return String.format(Locale.ROOT, "%.2f", v);
    }
}