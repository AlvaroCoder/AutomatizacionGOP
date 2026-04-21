package teucost.modelos;
import java.util.HashMap;
import java.util.List;

public class LecturaThroughput {
    private int filaVisita = 8;
    private String rutaThroughput;
    private String nombreHoja;
    private String nroVisita;

    public LecturaThroughput(){}

    public LecturaThroughput(String rutaThroughput){
        this.rutaThroughput =  rutaThroughput;
    }

    public HashMap<String, Integer> extraerResumenMovTeus() {
        LecturaExcels lecturaThroughput = null;
        HashMap<String, Integer> resumenMovTeus = new HashMap<>();
        resumenMovTeus.put("sumaMovContenedores", 0);
        resumenMovTeus.put("sumaTeus", 0);

        try {
            lecturaThroughput = new LecturaExcels(this.rutaThroughput);
            List<String> filaVisitas = lecturaThroughput.leerFila(nombreHoja, filaVisita);

            int sumaMovContenedores = 0;
            int sumaTeus = 0;

            for (int i = 0; i < filaVisitas.size(); i++) {
                String visita = filaVisitas.get(i);
                if (!visita.trim().equals(this.nroVisita.trim())) continue;

                List<String> columnaVisita = lecturaThroughput.leerColumna(nombreHoja, i);

                for (int j = 30; j < 38; j++) {
                    if (j >= columnaVisita.size()) {
                        System.out.println("[WARN] Columna tiene menos filas de las esperadas en j=" + j);
                        break;
                    }
                    String valorStr = columnaVisita.get(j).trim();
                    if (!valorStr.isEmpty()) {
                        sumaMovContenedores += (int) Double.parseDouble(valorStr.replace(",", "."));
                    }
                }

                String teusStr = lecturaThroughput.leerCelda(nombreHoja, 39, i).trim();
                if (!teusStr.isEmpty()) {
                    sumaTeus = (int) Double.parseDouble(teusStr.replace(",", "."));
                }
            }

            resumenMovTeus.put("sumaMovContenedores", sumaMovContenedores);
            resumenMovTeus.put("sumaTeus", sumaTeus);

        } catch (Exception err) {
            System.out.println("[ERROR] extraerResumenMovTeus: " + err);
            err.printStackTrace();
            // ✅ Bug 1 corregido: retorna mapa con valores en 0 en lugar de no retornar nada
        } finally {
            if (lecturaThroughput != null) {
                try { lecturaThroughput.cerrar(); } catch (Exception ignored) {}
            }
        }

        // ✅ Un solo punto de retorno — siempre retorna, con o sin error
        return resumenMovTeus;
    }

    public void setRutaThroughput(String rutaThroughput) {
        this.rutaThroughput = rutaThroughput;
    }

    public void setNroVisita(String nroVisita) {
        this.nroVisita = nroVisita;
    }

    public void setNombreHoja(String nombreHoja) {
        this.nombreHoja = nombreHoja;
    }

    public void setFilaVisita(int filaVisita) {
        this.filaVisita = filaVisita;
    }

}
