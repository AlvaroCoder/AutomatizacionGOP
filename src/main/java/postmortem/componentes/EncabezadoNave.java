package postmortem.componentes;

public class EncabezadoNave implements Componente {
    private final String nombreNave;
    private final String fechaNave;
    private final String linea;
    private final String servicio;
    private final String visita;
    private final String muelle;
    private final String inicioOperaciones;
    private final String finOperaciones;

    public EncabezadoNave(String nombreNave, String fechaNave,
                          String linea,     String servicio,
                          String visita,    String muelle,
                          String inicioOperaciones, String finOperaciones) {
        this.nombreNave        = nombreNave;
        this.fechaNave         = fechaNave;
        this.linea             = linea;
        this.servicio          = servicio;
        this.visita            = visita;
        this.muelle            = muelle;
        this.inicioOperaciones = inicioOperaciones;
        this.finOperaciones    = finOperaciones;
    }

    @Override
    public String render(){
        return "<div class='ctn-header-nave'>"
                    + celdaNave()
                    + "<div class='ctn-datos-nave'>"
                        + celda(linea, "Linea")
                        + celda(servicio, "Servicio")
                        + celda(visita, "Visita")
                        + celda(muelle, "Muelle")
                        + celda(inicioOperaciones, "Inicio de Operaciones")
                        + celda(finOperaciones, "Fin de Operaciones")
                    + "</div>" +
                "</div>";
    }

    private String celdaNave(){
        return "<div class='nav-nave'>" +
                "<h1 class='title-nave'>"+ nombreNave +"</h1>" +
                "<p class='text-nave'>"+ fechaNave +"</p>" +
                "</div>";
    }

    private String celda(String valor, String etiqueta){
        return "<div class='item-datos'>" +
                "<h1 class='title-datos'>"+valor+"</h1>" +
                "<p class='text-datos'>"+etiqueta+"</p>" +
                "</div>";
    }

    @Override
    public String css(){
        return ".ctn-header-nave {\n" +
                "  flex: 1;\n" +
                "  background-color: #e8f3ff;\n" +
                "  padding-left: 8px;\n" +
                "  padding-right: 8px;\n" +
                "\n" +
                "  display: flex;\n" +
                "  flex-direction: row;\n" +
                "  justify-content: space-between;\n" +
                "}\n" +
                "\n" +
                ".ctn-datos-nave {\n" +
                "  flex: 1;\n" +
                "  display: flex;\n" +
                "  flex-direction: row;\n" +
                "  justify-content: space-between;\n" +
                "  align-items: center;\n" +
                "  gap: 10px;\n" +
                "  margin-left: 8px;\n" +
                "}\n" +
                "\n" +
                ".item-datos {\n" +
                "  border-left: 2px solid #213a7c;\n" +
                "  padding-left: 5px;\n" +
                "  padding-right: 5px;\n" +
                "}\n" +
                "\n" +
                ".title-datos {\n" +
                "  font-size: medium;\n" +
                "  color: #252423;\n" +
                "  margin-bottom: 3px;\n" +
                "}\n" +
                ".text-datos {\n" +
                "  margin-top: 0px;\n" +
                "  color: #213a7c;\n" +
                "}\n" +
                "\n" +
                ".nav-nave {\n" +
                "  width: fit-content;\n" +
                "  text-align: center;\n" +
                "  margin-top: 0px;\n" +
                "  padding: 0px;\n" +
                "}\n" +
                "\n" +
                ".title-nave {\n" +
                "  font-size: large;\n" +
                "  color: #597fb9;\n" +
                "  border: 1px solid #213a7c;\n" +
                "  background-color: white;\n" +
                "  margin-bottom: 0px;\n" +
                "  padding: 8px;\n" +
                "  border-radius: 20px 20px 0px 0px;\n" +
                "  margin-top: 0px;\n" +
                "}\n" +
                "\n" +
                ".text-nave {\n" +
                "  margin-top: 0px;\n" +
                "  border: 1px solid #213a7c;\n" +
                "  padding: 5px;\n" +
                "  border-radius: 0px 0px 20px 20px;\n" +
                "  margin-bottom: 0px;\n" +
                "}";
    }
}
