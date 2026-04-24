package postmortem.controlador;

import postmortem.modelo.ExcelMoveHistoryPostMortem;
import postmortem.modelo.ExcelNavesPostMortem;
import java.util.List;

public class LecturaKPIsPostMortem {

    private List<ExcelMoveHistoryPostMortem.DatosMoveHistory> excelMoveHistory;
    private List<ExcelNavesPostMortem.DatosNave> excelNaves;

    private LecturaMoveHistoryPostMortem lecturaMoveHistory;
    private LecturaNaves lecturaNaves;

    public LecturaKPIsPostMortem(String rutaExcelNaves, String rutaExcelMoveHistory){
        lecturaNaves = new LecturaNaves(rutaExcelNaves);
        lecturaMoveHistory = new LecturaMoveHistoryPostMortem(rutaExcelMoveHistory);

        excelMoveHistory = lecturaMoveHistory.extraerDatosMoveHistory();
        excelNaves = lecturaNaves.extraerDatosNave();
    }

    public void calcularKpis(){
        List<ExcelMoveHistoryPostMortem.DatosMoveHistory>
                datosMoveHistory = lecturaMoveHistory.extraerDatosMoveHistory();

        List<ExcelNavesPostMortem.DatosNave>
                datosNaves = lecturaNaves.extraerDatosNave();

        for (ExcelNavesPostMortem.DatosNave nave : datosNaves){
            System.out.println("Nave "+nave.getVisita());
        }

    }


}
