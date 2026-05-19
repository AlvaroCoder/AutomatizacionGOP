package teucost.modelos.excels;

public class ExcelConciliado {

    private String rutaExcel;

    // Columnas con valores por defecto — se pueden sobreescribir
    // si el encabezado cambia de posición
    private int colVisita = 3;
    private int colMuelle = 5;

    public ExcelConciliado(String rutaExcel) {
        this.rutaExcel = rutaExcel;
    }

    public String getRutaExcel() { return rutaExcel; }
    public int getColVisita()    { return colVisita; }
    public int getColMuelle()    { return colMuelle; }
    public void setColVisita(int colVisita) { this.colVisita = colVisita; }
    public void setColMuelle(int colMuelle) { this.colMuelle = colMuelle; }

    // ── Registro por fila ─────────────────────────────────────
    public static class DatosConciliado {
        public final String visita;
        public final String muelle;

        public DatosConciliado(String visita, String muelle) {
            this.visita = visita;
            this.muelle = muelle;
        }

        @Override
        public String toString() {
            return String.format("[Visita=%s | Muelle=%s]", visita, muelle);
        }
    }
}