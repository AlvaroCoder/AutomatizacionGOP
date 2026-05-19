package teucost.modelos.excels;

public class ExcelThroughput {

    private String rutaExcel;

    // Fila donde están los números de visita (encabezado de columnas)
    private int filaVisita = 8;

    // Rango de filas donde están los movimientos de contenedores
    private int filaMovIni = 30;
    private int filaMovFin = 37;   // inclusive

    // Fila donde está el total de TEUs
    private int filaTeus   = 39;

    public ExcelThroughput(String rutaExcel) {
        this.rutaExcel = rutaExcel;
    }

    public String getRutaExcel()  { return rutaExcel;  }
    public int getFilaVisita()    { return filaVisita;  }
    public int getFilaMovIni()    { return filaMovIni;  }
    public int getFilaMovFin()    { return filaMovFin;  }
    public int getFilaTeus()      { return filaTeus;    }

    public void setFilaVisita(int filaVisita) { this.filaVisita = filaVisita; }
    public void setFilaMovIni(int filaMovIni) { this.filaMovIni = filaMovIni; }
    public void setFilaMovFin(int filaMovFin) { this.filaMovFin = filaMovFin; }
    public void setFilaTeus(int filaTeus)     { this.filaTeus   = filaTeus;   }

    // ── Registro por visita ───────────────────────────────────
    public static class DatosThroughput {
        public final String visita;
        public final int    movContenedores;
        public final int    teus;

        public DatosThroughput(String visita, int movContenedores, int teus) {
            this.visita          = visita;
            this.movContenedores = movContenedores;
            this.teus            = teus;
        }

        @Override
        public String toString() {
            return String.format("[Visita=%s | MovCont=%d | TEUs=%d]",
                    visita, movContenedores, teus);
        }
    }
}