package teucost.modelos.excels;

public class ExcelDatosNave {
    private String rutaExcel;

    private final String keyColVesselName = "Vessel Name";
    private final String keyColLine = "Line";
    private final String keyColService = "Service";
    private final String keyColStartWork = "Start Work";
    private final String keyColEndWork = "End Work";
    private final String keyColCuadrillas = "CUADRILLAS";
    private final String keyColVisit = "Visit";

    private int colVesselName;
    private int colLine, colService, colVisit;

    private int colStartWork, colEndWork;
    private int colCuadrillas;

    public ExcelDatosNave(String rutaExcel){this.rutaExcel = rutaExcel;}

    public String getRutaExcel() {
        return rutaExcel;
    }

    public String getKeyColVesselName() {
        return keyColVesselName;
    }

    public String getKeyColLine() {
        return keyColLine;
    }

    public String getKeyColService() {
        return keyColService;
    }

    public String getKeyColStartWork() {
        return keyColStartWork;
    }

    public String getKeyColEndWork() {
        return keyColEndWork;
    }

    public String getKeyColCuadrillas() {
        return keyColCuadrillas;
    }

    public String getKeyColVisit() {
        return keyColVisit;
    }

    public int getColVesselName() {
        return colVesselName;
    }

    public int getColLine() {
        return colLine;
    }

    public int getColService() {
        return colService;
    }

    public int getColVisit() {
        return colVisit;
    }

    public int getColStartWork() {
        return colStartWork;
    }

    public int getColEndWork() {
        return colEndWork;
    }

    public int getColCuadrillas() {
        return colCuadrillas;
    }

    public void setRutaExcel(String rutaExcel) {
        this.rutaExcel = rutaExcel;
    }

    public void setColVesselName(int colVesselName) {
        this.colVesselName = colVesselName;
    }

    public void setColLine(int colLine) {
        this.colLine = colLine;
    }

    public void setColService(int colService) {
        this.colService = colService;
    }

    public void setColVisit(int colVisit) {
        this.colVisit = colVisit;
    }

    public void setColStartWork(int colStartWork) {
        this.colStartWork = colStartWork;
    }

    public void setColEndWork(int colEndWork) {
        this.colEndWork = colEndWork;
    }

    public void setColCuadrillas(int colCuadrillas) {
        this.colCuadrillas = colCuadrillas;
    }

    public static class DatosNave{
        public final String lineVessel;
        public final String vesselName;
        public final String visita;
        public final String service;
        public final String startWork;
        public final String endWork;
        public final String cuadrillas;

        public DatosNave(
                String visita,
                String vesselName,
                String lineVessel,
                String service,
                String startWork,
                String endWork,
                String cuadrillas
        ){
            this.vesselName = vesselName;
            this.visita = visita;
            this.lineVessel = lineVessel;
            this.service = service;
            this.startWork = startWork;
            this.endWork = endWork;
            this.cuadrillas = cuadrillas;
        }

        public String getLineVessel() {
            return lineVessel;
        }

        public String getVesselName() {
            return vesselName;
        }

        public String getVisita() {
            return visita;
        }

        public String getService() {
            return service;
        }

        public String getStartWork() {
            return startWork;
        }

        public String getEndWork() {
            return endWork;
        }

        public String getCuadrillas() {
            return cuadrillas;
        }

        @Override
        public String toString(){
            return String.format(
              "[Vessel Name=%s | Visita=%s | Service =%s | StartWork=%s | EndWork=%s]",
                      vesselName, visita, service, startWork, endWork
            );
        }
    }
}
