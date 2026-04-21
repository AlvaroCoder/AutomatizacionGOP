package teucost.modelos;

import java.util.List;

public class LecturaConciliado {

    private String rutaConciliado;
    private String nombreHoja = "Cont-1-26";

    private final int  colMuelle = 5;
    private final int  colVisita = 3;

    public LecturaConciliado(String rutaConciliado){
        this.rutaConciliado = rutaConciliado;
    }

    public String extraerMuelleNave(String visitaNave){
        String muelleNave = "";
        LecturaExcels lecturaConciliado = null;
        try {
            lecturaConciliado = new LecturaExcels(this.rutaConciliado);
            List<String> listaVisitas = lecturaConciliado.leerColumna(nombreHoja, colVisita);
            for (int i = 0; i < listaVisitas.size(); i++) {
                String visita = listaVisitas.get(i);
                if (visita.isEmpty() || visita.equals("V.V.")) continue;

                if (visita.equals(visitaNave)) {
                    // ✅ i es directamente el índice de fila real en el Excel
                    muelleNave = lecturaConciliado.leerCelda(nombreHoja, i+1, colMuelle);
                    System.out.println("[INFO] Visita '" + visitaNave
                            + "' encontrada en fila " + i+1
                            + " → Muelle: " + muelleNave);
                    break; // ✅ Detener al encontrar
                }
            }

        } catch (Exception err){
            System.out.println("err = " + err.getMessage());
        } finally {
            if (lecturaConciliado != null) {
                try { lecturaConciliado.cerrar(); } catch (Exception ignored) {}
            }
        }
        return muelleNave;
    }

}
