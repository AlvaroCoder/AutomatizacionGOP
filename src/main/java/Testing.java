import teucost.controladores.ControladorCostoTeu;

import java.util.Arrays;

public class Testing {
    public static void main(String[] args) {
        ControladorCostoTeu costoTeu = new ControladorCostoTeu(
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA THROUGHPUTS\\THROUGHPUT GOP 2026-5.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA NAVES\\DATA_NAVES_MAYO.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA CONCILIADOS\\Conciliado Trafico 2Q 2026 MAYO.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA OMS\\OM- May26.xlsm",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenMayo.xlsx"
        );

        costoTeu.setNombreHojaCosmos("DatosNave");

        costoTeu.setNumMesOM(5);
        costoTeu.setRutaExcelDestino("C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\RESUMEN EXCEL TEU COST.xlsx");
   costoTeu.procesarVisitas(Arrays.asList("227-26", "235-26", "250-26", "251-26", "252-26", "253-26", "254-26", "255-26", "256-26", "257-26", "258-26", "259-26", "260-26", "261-26", "262-26", "263-26", "264-26", "267-26", "268-26", "269-26", "270-26", "271-26", "272-26", "273-26", "274-26", "275-26", "276-26", "277-26", "278-26", "279-26", "280-26", "281-26", "282-26", "283-26", "284-26", "285-26", "286-26", "287-26", "288-26", "289-26", "290-26", "291-26", "293-26", "294-26", "295-26", "296-26", "297-26"
       ));
        // "180-26", "187-26", "190-26", "191-26", "192-26", "193-26", "194-26", "195-26", "196-26", "197-26", "198-26", "199-26", "200-26", "201-26", "202-26", "203-26", "204-26", "205-26", "206-26", "207-26", "208-26", "209-26", "210-26", "212-26", "213-26", "214-26", "215-26", "216-26", "217-26", "218-26", "219-26", "221-26", "222-26", "223-26", "225-26", "226-26", "228-26", "229-26", "230-26", "231-26", "232-26", "233-26", "234-26", "245-26"
    }

}
