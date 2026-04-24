package postmortem.modelo;

public class ExcelNavesPostMortem {

    private String rutaExcel;
    private int colVisita, colLinea, colServicio;
    private int colInicioOperaciones, colFinOperaciones;

    private final String keyColInicio   = "Start Work";
    private final String keyColFin      = "End Work";
    private final String keyColVisita   = "Visit";
    private final String keyColLinea    = "Line";
    private final String keyColServicio = "Service";

    public ExcelNavesPostMortem(String rutaExcel) {
        this.rutaExcel = rutaExcel;
    }

    public String getKeyColInicio()   { return keyColInicio; }
    public String getKeyColFin()      { return keyColFin; }
    public String getKeyColVisita()   { return keyColVisita; }
    public String getKeyColLinea()    { return keyColLinea; }
    public String getKeyColServicio() { return keyColServicio; }
    public String getRutaExcel()      { return rutaExcel; }

    public int getColVisita()              { return colVisita; }
    public void setColVisita(int v)        { this.colVisita = v; }

    public int getColLinea()               { return colLinea; }
    public void setColLinea(int v)         { this.colLinea = v; }

    public int getColServicio()            { return colServicio; }
    public void setColServicio(int v)      { this.colServicio = v; }

    public int getColInicioOperaciones()          { return colInicioOperaciones; }
    public void setColInicioOperaciones(int v)    { this.colInicioOperaciones = v; }

    public int getColFinOperaciones()             { return colFinOperaciones; }
    public void setColFinOperaciones(int v)       { this.colFinOperaciones = v; }

    // ── POJO por fila de nave ────────────────────────────────
    public static class DatosNave {
        public final String visita;
        public final String linea;
        public final String servicio;
        public final String inicioOperaciones;
        public final String finOperaciones;

        public DatosNave(String visita, String linea, String servicio,
                         String inicioOperaciones, String finOperaciones) {
            this.visita             = visita;
            this.linea              = linea;
            this.servicio           = servicio;
            this.inicioOperaciones  = inicioOperaciones;
            this.finOperaciones     = finOperaciones;
        }

        public String getVisita() {
            return visita;
        }

        public String getLinea() {
            return linea;
        }

        public String getServicio() {
            return servicio;
        }

        public String getInicioOperaciones() {
            return inicioOperaciones;
        }

        public String getFinOperaciones() {
            return finOperaciones;
        }

        @Override
        public String toString() {
            return String.format(
                    "[Visita=%s | Linea=%s | Servicio=%s | Inicio=%s | Fin=%s]",
                    visita, linea, servicio, inicioOperaciones, finOperaciones
            );
        }
    }
}