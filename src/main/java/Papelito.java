import nombradas.AgruparNombradas;
import postmortem.controlador.LecturaKPIsPostMortem;
import postmortem.controlador.LecturaNaves;
import postmortem.modelo.ExcelNavesPostMortem;

public class Papelito {
    public static void main(String[] args) {
        LecturaKPIsPostMortem kpIsPostMortem = new LecturaKPIsPostMortem(
                "C:\\Users\\alvaro.pupuche\\Documents\\DataPostMortem\\VisitasAbril.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenMarzo.xlsx"
        );
        kpIsPostMortem.calcularKpis();
    }
}
