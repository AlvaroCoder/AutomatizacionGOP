import nombradas.AgruparNombradas;
import postmortem.controlador.LecturaKPIsPostMortem;
import postmortem.controlador.LecturaNaves;
import postmortem.modelo.ExcelNavesPostMortem;
import postmortem.pdf.PaginaIndicadores;

import java.util.List;
import java.util.Map;

public class Papelito {
    public void generarPaginas(List<Map<String, Object>> dataKpisNave){
        String rutaBase = "C:\\Users\\alvaro.pupuche\\Documents\\POST MORTEMS - PAGINAS";
        PaginaIndicadores pagina = new PaginaIndicadores();

        for (Map<String, Object> nave : dataKpisNave){
            String visita = nave.get("Visita").toString();
            String rutaPdf = rutaBase+"\\postmortem_"+visita+".pdf";

            pagina.generarPagina(nave, rutaPdf);
        }
    }
    public static void main(String[] args) {
        Papelito papelito = new Papelito();

        LecturaKPIsPostMortem kpIsPostMortem = new LecturaKPIsPostMortem(
                "C:\\Users\\alvaro.pupuche\\Documents\\DataPostMortem\\VisitasAbril.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenABR.xlsx"
        );
        List<Map<String, Object>> dataKpisNaves = kpIsPostMortem.calcularKpis();
        Map<String, Object> navePrueba = dataKpisNaves.get(0);

        PaginaIndicadores pagina = new PaginaIndicadores();
        String visita = navePrueba.get("Visita").toString();
        String rutaPdf = "C:\\Users\\alvaro.pupuche\\Documents\\POST MORTEMS - PAGINAS\\postmortem_"+visita+".pdf";

        pagina.generarPagina(navePrueba, rutaPdf);

       // papelito.generarPaginas(dataKpisNaves);
    }
}
