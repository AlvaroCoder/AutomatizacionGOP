package teucost.modelos;

import java.util.HashMap;
import java.util.List;

public class LecturaOmMensual {
    private String rutaOmMensual;

    private String nombreHojaOMDiesel;
    private String nombreHojaEnergia;
    private String nombreHojaEstibas;

    private LecturaExcels lecturaOM = null;
    private int numMes;

    private final int filaCostoMovRTG = 62;
    private final int filaCostoMovSTK = 32;
    private final int filaCostoMovEH  = 44;
    private final int filaCostoMovTT  = 98;
    private final int filaCostoUnitGalon = 96;
    private final int filaCostoGMGalonHora = 91;

    private final int filaRNombre = 25;

    private final int filaCosto28 = 27;
    private final int filaCosto39 = 38;
    private final int filaCosto48 = 47;

    private final int filaCosto29 = 28;
    private final int filaCosto40 = 39;
    private final int filaCosto49 = 48;

    private final int filaCosto57 = 56;
    private final int filaCosto66 = 65;

    private final int filaCosto58 = 57;
    private final int filaCosto67 = 66;

    public LecturaOmMensual(String rutaOmMensual) {
        this.rutaOmMensual = rutaOmMensual;
        try {
            this.lecturaOM = new LecturaExcels(rutaOmMensual);
        } catch (Exception err) {
            System.out.println("[ERROR] No se pudo abrir el archivo: " + err.getMessage());
        }
    }

    public int extraerColummnaReal(int numFila, String hojaSeleccionada) {
        int colRealSeleccionada = -1; // ✅ Bug 2: usar -1 como centinela en lugar de 0
        try {
            List<String> listaRealData = lecturaOM.leerFila(hojaSeleccionada, numFila);
            String textoBuscado = "R26-" + numMes;

            for (int i = 0; i < listaRealData.size(); i++) {
                if (listaRealData.get(i).trim().equals(textoBuscado)) {
                    colRealSeleccionada = i;
                    break; // ✅ Detener al encontrar
                }
            }

            if (colRealSeleccionada == -1) {
                System.out.println("[WARN] No se encontró columna para: '" + textoBuscado + "'");
            } else {
                System.out.println("[INFO] Columna encontrada para '"
                        + textoBuscado + "': " + colRealSeleccionada);
            }

        } catch (Exception err) {
            System.out.println("[ERROR] extraerColummnaReal: " + err.getMessage());
        }
        return colRealSeleccionada;
    }

    public HashMap<String, Double> extraerResumenCostosDiesel() {
        HashMap<String, Double> resumenCostos = new HashMap<>();
        resumenCostos.put("costoPorMovRTG", 0.0);
        resumenCostos.put("costoPorMovSTK", 0.0);
        resumenCostos.put("costoPorMovEH",  0.0);
        resumenCostos.put("costoPorMovTT",  0.0);
        resumenCostos.put("costoUnitGalon", 0.0);
        resumenCostos.put("costoGMGalonHora", 0.0);
        try {
            // ✅ Bug 2 corregido: validar que la columna fue encontrada
            int colSeleccionada = extraerColummnaReal(2, nombreHojaOMDiesel);
            if (colSeleccionada == -1) {
                System.out.println("[ERROR] No se puede extraer costos: columna no encontrada.");
                return resumenCostos;
            }

            // ✅ Bug 3 corregido: parsear con protección
            resumenCostos.put("costoPorMovRTG",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoMovRTG, colSeleccionada)));
            resumenCostos.put("costoPorMovSTK",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoMovSTK, colSeleccionada)));
            resumenCostos.put("costoPorMovEH",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoMovEH,  colSeleccionada)));
            resumenCostos.put("costoPorMovTT",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoMovTT,  colSeleccionada)));
            resumenCostos.put("costoUnitGalon",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoUnitGalon, colSeleccionada)));
            resumenCostos.put("costoGMGalonHora",
                    parsearDouble(lecturaOM.leerCelda(nombreHojaOMDiesel, filaCostoGMGalonHora, colSeleccionada)));

        } catch (Exception err) {
            System.out.println("[ERROR] extraerResumenCostos: " + err.getMessage());
            err.printStackTrace();
        } finally {
            // ✅ Bug 1 corregido: cerrar aquí, al final de todo el proceso
            if (lecturaOM != null) {
                try { lecturaOM.cerrar(); } catch (Exception ignored) {}
            }
        }

        return resumenCostos;
    }

    public HashMap<String, Double> extraerResumenCostosEnergia(){
        HashMap <String, Double> resumenCostos = new HashMap<>();
        resumenCostos.put("costoSTSPorHora", 0.0);
        resumenCostos.put("costoeRTGPorHora", 0.0);
        try {
            int colSeleccionada = extraerColummnaReal(25, nombreHojaEnergia);
            if (colSeleccionada == -1) {
                System.out.println("[ERROR] No se puede extraer costos: columna no encontrada.");
                return resumenCostos;
            }

            double costoCelda28 = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto28, colSeleccionada));
            double horaCelda29  = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto29, colSeleccionada));
            double costoEnergiaSTS01 = horaCelda29 != 0 ? (costoCelda28 * 1000) / horaCelda29 : 0.0;

            double costoCelda39 = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto39, colSeleccionada));
            double horaCelda40  = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto40, colSeleccionada));
            double costoEnergiaSTS02 = horaCelda40 != 0 ? (costoCelda39 * 1000) / horaCelda40 : 0.0;

            double costoCelda50 = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto48, colSeleccionada));
            double horaCelda51  = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto49, colSeleccionada));
            double costoEnergiaSTS03 = horaCelda51 != 0 ? (costoCelda50 * 1000) / horaCelda51 : 0.0;

            double costoCelda61 = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto57, colSeleccionada));
            double horaCelda62  = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto58, colSeleccionada));
            double costoEnergiaRTG05 = horaCelda62 != 0 ? (costoCelda61 * 1000) / horaCelda62 : 0.0;

            double costoCelda72 = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto66, colSeleccionada));
            double horaCelda73  = parsearDouble(lecturaOM.leerCelda(nombreHojaEnergia, filaCosto67, colSeleccionada));
            double costoEnergiaRTG06 = horaCelda73 != 0 ? (costoCelda72 * 1000) / horaCelda73 : 0.0;

            resumenCostos.put("costoEnergiaSTS01", costoEnergiaSTS01);
            resumenCostos.put("costoEnergiaSTS02", costoEnergiaSTS02);
            resumenCostos.put("costoEnergiaSTS03", costoEnergiaSTS03);
            resumenCostos.put("costoEnergiaRTG05", costoEnergiaRTG05);
            resumenCostos.put("costoEnergiaRTG06", costoEnergiaRTG06);
        } catch (Exception err) {
            System.out.println("[ERROR] extraerResumenCostos: " + err.getMessage());
            err.printStackTrace();
        } finally {
            // ✅ Bug 1 corregido: cerrar aquí, al final de todo el proceso
            if (lecturaOM != null) {
                try { lecturaOM.cerrar(); } catch (Exception ignored) {}
            }
        }
        return resumenCostos;
    }

    public double extraerCostoEstiba(){
        double valorCostoEstiba = 0.0;

        try {
            int colSeleccionada = extraerColummnaReal(2, nombreHojaEstibas);
            valorCostoEstiba = parsearDouble(lecturaOM.leerCelda(nombreHojaEstibas, 15, colSeleccionada));

        } catch (Exception err){
            System.out.println("err = " + err.getMessage());
        }finally {
            // ✅ Bug 1 corregido: cerrar aquí, al final de todo el proceso
            if (lecturaOM != null) {
                try { lecturaOM.cerrar(); } catch (Exception ignored) {}
            }
        }

        return valorCostoEstiba;
    }

    private double parsearDouble(String valor) {
        if (valor == null || valor.trim().isEmpty()) return 0.0;
        try {
            return Double.parseDouble(valor.trim().replace(",", "."));
        } catch (NumberFormatException e) {
            System.out.println("[WARN] No se pudo parsear: '" + valor + "'");
            return 0.0;
        }
    }

    public void setNumMes(int numMes)           { this.numMes = numMes; }
    public void setNombreHojaOMDiesel(String nombre)  { this.nombreHojaOMDiesel = nombre; }
    public void setNombreHojaEnergia(String nombre) {this.nombreHojaEnergia = nombre;}
    public void setNombreHojaEstibas(String nombreHojaEstibas) {this.nombreHojaEstibas = nombreHojaEstibas;}

}