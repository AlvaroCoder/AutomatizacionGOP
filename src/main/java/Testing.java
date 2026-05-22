import teucost.controladores.ControladorCostoTeu;

import java.util.Arrays;

public class Testing {
    public static void main(String[] args) {
        ControladorCostoTeu costoTeu = new ControladorCostoTeu(
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA THROUGHPUTS\\THROUGHPUT GOP 2026 4.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA NAVES\\DATA_NAVES_ABRIL.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA CONCILIADOS\\Conciliado Trafico 2Q 2026 ABRIL.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA OMS\\OM- ABRIL 2026.xlsm",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenABR.xlsx"
        );

        costoTeu.setNombreHojaCosmos("DatosNave");
        costoTeu.setNumMesOM(4);
        costoTeu.setRutaExcelDestino("C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\RESUMEN EXCEL TEU COST.xlsx");
        costoTeu.procesarVisitas(Arrays.asList("180-26", "187-26", "190-26", "191-26", "192-26"));
        // "180-26", "187-26", "190-26", "191-26", "192-26", "193-26", "194-26", "195-26", "196-26", "197-26", "198-26", "199-26", "200-26", "201-26", "202-26", "203-26", "204-26", "205-26", "206-26", "207-26", "208-26", "209-26", "210-26", "212-26", "213-26", "214-26", "215-26", "216-26", "217-26", "218-26", "219-26", "221-26", "222-26", "223-26", "225-26", "226-26", "228-26", "229-26", "230-26", "231-26", "232-26", "233-26", "234-26", "245-26"
    }

}
