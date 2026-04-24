import nombradas.AgruparNombradas;
import postmortem.controlador.LecturaKPIsPostMortem;
import postmortem.controlador.LecturaNaves;
import postmortem.modelo.ExcelNavesPostMortem;
import postmortem.pdf.PaginaIndicadores;

import java.util.List;
import java.util.Map;

public class Papelito {
    public static void main(String[] args) {
        LecturaKPIsPostMortem kpIsPostMortem = new LecturaKPIsPostMortem(
                "C:\\Users\\alvaro.pupuche\\Documents\\DataPostMortem\\VisitasAbril.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenABR.xlsx"
        );
        List<Map<String, Object>> dataKpisNaves = kpIsPostMortem.calcularKpis();
        Map<String, Object> navePrueba = dataKpisNaves.get(0);

        PaginaIndicadores pagina = new PaginaIndicadores();
        String visita = navePrueba.get("Visita").toString();
        String rutaPdf = "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\postmortem_"+visita+".pdf";

        pagina.generarPagina(navePrueba, rutaPdf);


    }
}
