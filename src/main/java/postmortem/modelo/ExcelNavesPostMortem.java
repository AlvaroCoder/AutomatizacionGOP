package postmortem.modelo;

public class ExcelNavesPostMortem {

    private String rutaExcel;
    private int colVisita, colLinea, colServicio;
    private int colInicioOperaciones, colFinOperaciones;
    private int colBerthATA, colBerthATD;

    private final String keyColInicio = "Start Work";
    private final String keyColFin = "End Work";
    private final String keyColVisita = "Visit";
    private final String keyColLinea = "Line";
    private final String keyColServicio = "Service";
    private final String keyColBerthATA = "ATA";
    private final String keyColBerthATD = "ATD";

    public ExcelNavesPostMortem(String rutaExcel) {
        this.rutaExcel = rutaExcel;
    }

    public String getKeyColInicio()   { return keyColInicio; }
    public String getKeyColFin()      { return keyColFin; }
    public String getKeyColVisita()   { return keyColVisita; }
    public String getKeyColLinea()    { return keyColLinea; }
    public String getKeyColServicio() { return keyColServicio; }
    public String getRutaExcel()      { return rutaExcel; }
    public String getKeyColBerthATA() {return keyColBerthATA;}
    public String getKeyColBerthATD() {return keyColBerthATD;}

    public int getColVisita()              { return colVisita; }
    public void setColVisita(int v)        { this.colVisita = v; }

    public int getColLinea()               { return colLinea; }
    public void setColLinea(int v)         { this.colLinea = v; }

    public int getColServicio()            { return colServicio; }
    public void setColServicio(int v)      { this.colServicio = v; }

    public int getColInicioOperaciones()          { return colInicioOperaciones; }
    public void setColInicioOperaciones(int v) { this.colInicioOperaciones = v; }

    public int getColFinOperaciones() { return colFinOperaciones; }
    public void setColFinOperaciones(int v) { this.colFinOperaciones = v; }

    public int getColBerthATA(){return  this.colBerthATA;}
    public void setColBerthATA(int v){this.colBerthATA = v;}

    public int getColBerthATD(){return  this.colBerthATD;}
    public void setColBerthATD(int v){this.colBerthATD = v;}

    // ── POJO por fila de nave ────────────────────────────────
    public static class DatosNave {
        public final String visita;
        public final String linea;
        public final String servicio;
        public final String inicioOperaciones;
        public final String finOperaciones;
        public final String berthATA;
        public final String berthATD;


        public DatosNave(String visita, String linea, String servicio,
                         String inicioOperaciones, String finOperaciones,
                         String berthATA, String berthATD) {
            this.visita             = visita;
            this.linea              = linea;
            this.servicio           = servicio;
            this.inicioOperaciones  = inicioOperaciones;
            this.finOperaciones     = finOperaciones;
            this.berthATA = berthATA;
            this.berthATD = berthATD;
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

        public String getBerthATA(){return  berthATA;}

        public String getBerthATD(){return berthATD;}


        @Override
        public String toString() {
            return String.format(
                    "[Visita=%s | Linea=%s | Servicio=%s | Inicio=%s | Fin=%s | BerthATA=%s | BerthATD=%s]",
                    visita, linea, servicio, inicioOperaciones, finOperaciones, berthATA, berthATD
            );
        }
    }
}