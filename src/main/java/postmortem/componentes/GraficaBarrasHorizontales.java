package postmortem.componentes;

import java.util.List;

public class GraficaBarrasHorizontales implements Componente {

    public static class Barra {
        public final String label;
        public final double valor;
        public final String color;
        public final String badgeTexto;   // null si no hay badge

        public Barra(String label, double valor, String color, String badgeTexto) {
            this.label      = label;
            this.valor      = valor;
            this.color      = color;
            this.badgeTexto = badgeTexto;
        }
    }

    private final String        titulo;
    private final String        totalLabel;
    private final List<Barra>   barras;
    private final String        unidad;
    private final String        cssId;   // prefijo único para no colisionar CSS

    public GraficaBarrasHorizontales(String titulo, String totalLabel,
                                     List<Barra> barras, String unidad,
                                     String cssId) {
        this.titulo     = titulo;
        this.totalLabel = totalLabel;
        this.barras     = barras;
        this.unidad     = unidad;
        this.cssId      = cssId;
    }

    @Override
    public String render() {
        double maxValor = barras.stream()
                .mapToDouble(b -> b.valor)
                .max().orElse(1.0);

        StringBuilder sb = new StringBuilder();
        sb.append("<div class='chart-section-" + cssId + "'>");
        sb.append("<div class='chart-title'>" + titulo + "</div>");
        sb.append("<div class='chart-total'>" + totalLabel + "</div>");
        sb.append("<div class='bar-chart-" + cssId + "'>");

        for (Barra barra : barras) {
            int pct = (int) ((barra.valor / maxValor) * 100);
            String valorFmt = barra.valor == Math.floor(barra.valor)
                    ? String.valueOf((int) barra.valor)
                    : String.format(java.util.Locale.ROOT, "%.2f", barra.valor);

            sb.append("<div class='bar-row-" + cssId + "'>"
                    + "<span class='bar-label-" + cssId + "'>" + barra.label + "</span>"
                    + "<div class='bar-track-" + cssId + "'>"
                    + "<div class='bar-fill-" + cssId + "' "
                    + "  style='width:" + pct + "%; background:" + barra.color + ";'>"
                    + "<span class='bar-value-" + cssId + "'>" + valorFmt
                    + (unidad.isEmpty() ? "" : " " + unidad) + "</span>"
                    + "</div></div>");

            if (barra.badgeTexto != null) {
                sb.append("<span class='bar-badge-" + cssId + "' "
                        + "style='background:" + barra.color + "'>"
                        + barra.badgeTexto + "</span>");
            }
            sb.append("</div>");
        }

        sb.append("</div></div>");
        return sb.toString();
    }

    @Override
    public String css() {
        String p = cssId; // alias corto
        return ".chart-section-" + p + " { margin-bottom: 20px; }"
                + ".chart-title  { font-weight: bold; font-size: 13px; "
                + "  text-align: center; margin-bottom: 4px; }"
                + ".chart-total  { text-align: center; font-size: 22px; "
                + "  font-weight: bold; margin-bottom: 10px; }"
                + ".bar-chart-"  + p + " { padding: 0 10px; }"
                + ".bar-row-"    + p + " { display: flex; align-items: center; "
                + "  margin-bottom: 8px; gap: 8px; }"
                + ".bar-label-"  + p + " { min-width: 120px; font-size: 10px; "
                + "  text-align: right; font-weight: bold; }"
                + ".bar-track-"  + p + " { flex: 1; background: #ddd; "
                + "  border-radius: 3px; height: 24px; }"
                + ".bar-fill-"   + p + " { height: 100%; border-radius: 3px; "
                + "  display: flex; align-items: center; "
                + "  justify-content: flex-end; padding-right: 6px; }"
                + ".bar-value-"  + p + " { color: white; font-weight: bold; font-size: 11px; }"
                + ".bar-badge-"  + p + " { color: white; padding: 2px 8px; "
                + "  border-radius: 3px; font-size: 10px; font-weight: bold; "
                + "  min-width: 35px; text-align: center; }";
    }
}