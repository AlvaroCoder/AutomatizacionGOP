import teucost.controladores.ControladorCostoTeu;

import java.util.Arrays;

public class TestCostoTeu   {
    public static void main(String[] args) {
        ControladorCostoTeu controladorCostoTeu = new ControladorCostoTeu(
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA THROUGHPUTS\\THROUGHPUT GOP 2026 4.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA NAVES\\DATA_NAVES_ABRIL.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA CONCILIADOS\\Conciliado Trafico 2Q 2026 ABRIL.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA OMS\\OM- ABRIL 2026.xlsm",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenABR.xlsx"
        );
        controladorCostoTeu.setRutaExcelDestino("C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\RESUMEN EXCEL TEU COST.xlsx");
        controladorCostoTeu.setNombreHojaCosmos("DatosNave");
        controladorCostoTeu.setNumMesOM(4);
        controladorCostoTeu.escribirVariasVisitasEnTabla(
                Arrays.asList("230-26", "231-26", "232-26", "233-26", "234-26", "245-26")
        );

    }
}
