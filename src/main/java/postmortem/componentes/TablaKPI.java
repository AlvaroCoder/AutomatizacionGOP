package postmortem.componentes;

import java.util.List;
import java.util.Map;

public class TablaKPI implements Componente {

    private final double craneIntensity;
    private final int    gruasUtilizadas;
    private final double duracionOp;
    private final int    longGang;
    private final double bmph;
    private final double gmph;
    private final Map<String, Object> kpi;

    public TablaKPI(double craneIntensity, int gruasUtilizadas,
                    double duracionOp,     int longGang,
                    double bmph,           double gmph,
                    Map<String, Object> kpi) {
        this.craneIntensity  = craneIntensity;
        this.gruasUtilizadas = gruasUtilizadas;
        this.duracionOp      = duracionOp;
        this.longGang        = longGang;
        this.bmph            = bmph;
        this.gmph            = gmph;
        this.kpi             = kpi;
    }

    @Override
    public String render() {
        return "<table class='kpi-table'>"
                + encabezadoTabla()
                + "<tbody>"
                + fila("Crane Intensity",    "", "QC's", fmt(craneIntensity),         "QC's")
                + fila("Grúas utilizadas",   "", "QC's", fmt((double) gruasUtilizadas),"QC's")
                + fila("Duración Operación", "", "Hrs",  fmt(duracionOp),             "Hrs")
                + fila("Long Gang",          "", "Movs", fmt((double) longGang),      "Movs")
                + fila("BMPH",               "", "MPH",  fmt(bmph),                   "MPH")
                + fila("GMPH",               "", "MPH",  fmt(gmph),                   "MPH")
                + filasMovimientos()
                + filasDemoras()
                + "</tbody></table>";
    }

    private String encabezadoTabla() {
        return "<thead><tr>"
                + "<th class='col-indicador'></th>"
                + "<th colspan='2' class='col-planeado'>Planeado</th>"
                + "<th colspan='2' class='col-ejecutado'>Ejecutado</th>"
                + "</tr></thead>";
    }

    private String fila(String indicador, String planVal, String planUnid,
                        String ejVal,    String ejUnid) {
        return "<tr class='kpi-light'>"
                + "<td class='td-indicador'>" + indicador + "</td>"
                + "<td class='td-val-plan'>"  + planVal   + "</td>"
                + "<td class='td-unid'>"      + planUnid  + "</td>"
                + "<td class='td-val-ejec'>"  + ejVal     + "</td>"
                + "<td class='td-unid'>"      + ejUnid    + "</td>"
                + "</tr>";
    }

    private String filasMovimientos() {
        int descarga    = getInt("Descarga");
        int embarque    = getInt("Embarque");
        int restibas    = getInt("Restibas");
        int transbordos = getInt("Transbordos");
        int total       = getInt("TotalMovimientos");

        return filaMovs("Descarga",    descarga,    true,  "FL-FM", "36 min")
                + filaMovs("Embarque",   embarque,    true,  "LM-LL", "44 min")
                + filaMovs("Restibas",   restibas,    false, null,    null)
                + filaMovs("Transbordos",transbordos, false, null,    null)
                + filaTotal(total);
    }

    private String filaMovs(String label, int val,
                            boolean conBadge, String badge, String min) {
        String extra = "";
        if (conBadge && badge != null) {
            extra = "<span class='badge-movs'>" + badge + "</span>"
                    + " <span class='badge-min'>" + min   + "</span>";
        }
        return "<tr class='kpi-movs'>"
                + "<td class='td-movs-label'>" + label + "</td>"
                + "<td class='td-movs-val'>"   + val   + "</td>"
                + "<td class='td-movs-unid'>movs</td>"
                + "<td colspan='2'>"           + extra + "</td>"
                + "</tr>";
    }

    private String filaTotal(int total) {
        return "<tr class='kpi-total'>"
                + "<td class='td-total-label'>MOVIMIENTOS</td>"
                + "<td class='td-total-val'>" + total + "</td>"
                + "<td class='td-total-unid'>movs</td>"
                + "<td></td><td></td>"
                + "</tr>";
    }

    @SuppressWarnings("unchecked")
    private String filasDemoras() {
        List<Map<String, Object>> demoras =
                (List<Map<String, Object>>) kpi.get("Demoras");
        if (demoras.isEmpty()) return "";

        StringBuilder sb = new StringBuilder();
        sb.append("<tr class='demoras-header'>"
                + "<td rowspan='" + (demoras.size() + 1) + "' class='td-demoras-titulo'>"
                + "<b>Demoras<br/>Principales</b></td>"
                + "<td>Demoras</td><td></td>"
                + "<td>Durac (Hrs)</td><td>% Hrs Qc</td></tr>");

        for (Map<String, Object> d : demoras) {
            sb.append("<tr class='demoras-row'>"
                    + "<td>" + d.getOrDefault("Nombre",   "") + "</td>"
                    + "<td></td>"
                    + "<td>" + fmt(getDbl(d, "DurHrs"))        + "</td>"
                    + "<td>" + fmt(getDbl(d, "PctHrsQc"))+ " %</td>"
                    + "</tr>");
        }
        return sb.toString();
    }

    @Override
    public String css() {
        return ".kpi-table { width: 100%; border-collapse: collapse; }"
                + ".kpi-table thead th { background: #1a2e4a; color: white; "
                + "  padding: 6px; text-align: center; }"
                + ".col-indicador { width: 38%; }"
                + ".col-planeado  { width: 31%; background: #5b7fa6 !important; }"
                + ".col-ejecutado { width: 31%; background: #2e4d7b !important; }"
                + ".kpi-light td  { padding: 5px 8px; border-bottom: 1px solid #ddd; "
                + "  background: #e8f0f8; }"
                + ".td-indicador  { text-align: left; }"
                + ".td-val-plan   { text-align: right; color: #aaa; }"
                + ".td-val-ejec   { text-align: right; font-weight: bold; font-size: 12px; }"
                + ".td-unid       { padding-left: 4px; color: #666; }"
                + ".kpi-movs td   { padding: 4px 8px; background: #2e4d7b; "
                + "  color: white; border-bottom: 1px solid #1a2e4a; }"
                + ".td-movs-val   { text-align: right; font-weight: bold; }"
                + ".td-movs-unid  { padding-left: 4px; }"
                + ".kpi-total td  { padding: 5px 8px; background: #1a2e4a; "
                + "  color: white; font-weight: bold; }"
                + ".td-total-val  { text-align: right; }"
                + ".badge-movs    { background: #1a2e4a; color: white; "
                + "  padding: 2px 6px; border-radius: 3px; font-weight: bold; }"
                + ".badge-min     { color: white; font-weight: bold; }"
                + ".demoras-header td { background: #2c2c2c; color: white; "
                + "  padding: 4px 8px; font-size: 10px; }"
                + ".td-demoras-titulo { background: #2c2c2c; color: white; "
                + "  text-align: center; vertical-align: middle; }"
                + ".demoras-row td { padding: 3px 8px; background: #f5f5f5; "
                + "  border-bottom: 1px solid #ddd; }";
    }

    // ── Helpers ───────────────────────────────────────────────
    private String fmt(double v) { return String.format(java.util.Locale.ROOT, "%.2f", v); }
    private int    getInt(String k) {
        Object v = kpi.get(k);
        return v instanceof Number ? ((Number) v).intValue() : 0;
    }
    private double getDbl(Map<String, Object> m, String k) {
        Object v = m.get(k);
        return v instanceof Number ? ((Number) v).doubleValue() : 0.0;
    }
}