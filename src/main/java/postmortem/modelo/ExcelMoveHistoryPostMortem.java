package postmortem.modelo;

public class ExcelMoveHistoryPostMortem {
    private String rutaExcel;
    private int colTimeCompleted;
    private int colCraneCheName, colPutCheName,
            colFetchCheName, colCarryCheName;

    private int colVisit;

    private int colFromPosition, colToPosition;

    private final String keyColVisita = "Carrier Visit";
    private final String keyColFetch = "Fetch CHE Name";
    private final String keyColPut = "Put CHE Name";
    private final String keyColCarry = "Carry CHE Name";
    private final String keyColTimeCompleted = "Time Completed";

    private final String keyColFromPosition = "From Position";
    private final String keyColToPosition = "To Position";

    public ExcelMoveHistoryPostMortem(String rutaExcel){
        this.rutaExcel = rutaExcel;
    }

    public String getRutaExcel() {return rutaExcel;}

    public String getKeyColTimeCompleted() {
        return keyColTimeCompleted;
    }

    public String getKeyColFromPosition() {
        return keyColFromPosition;
    }

    public String getKeyColToPosition() {
        return keyColToPosition;
    }

    public int getColTimeCompleted() {return colTimeCompleted;}

    public void setColTimeCompleted(int colTimeCompleted) {this.colTimeCompleted = colTimeCompleted;}

    public int getColCraneCheName() {return colCraneCheName;}

    public void setColCraneCheName(int colCraneCheName) {this.colCraneCheName = colCraneCheName;}

    public int getColPutCheName() {return colPutCheName;}

    public void setColPutCheName(int colPutCheName) {this.colPutCheName = colPutCheName;}

    public int getColFetchCheName() {return colFetchCheName;}

    public void setColFetchCheName(int colFetchCheName) {this.colFetchCheName = colFetchCheName;}

    public int getColCarryCheName() {return colCarryCheName;}

    public void setColCarryCheName(int colCarryCheName) {this.colCarryCheName = colCarryCheName;}

    public int getColVisit() {return colVisit;}

    public int getColFromPosition(){return colFromPosition;}

    public int getColToPosition(){return colToPosition;}

    public void setColFromPosition(int colFromPosition){this.colFromPosition = colFromPosition;}

    public void setColToPosition(int colToPosition){this.colToPosition = colToPosition;}

    public void setColVisit(int colVisit) {this.colVisit = colVisit;}

    public String getKeyColVisita() {return keyColVisita;}

    public String getKeyColFetch() {return keyColFetch;}

    public String getKeyColPut() {return keyColPut;}

    public String getKeyColCarry() {return keyColCarry;}

    public static class DatosMoveHistory{
        public final String fetchCheName;
        public final String putCheName;
        public final String carryCheName;
        public final String craneCheName;
        public final String visita;

        public final String fromPosition, toPosition;

        public DatosMoveHistory(
                String visita,
                String fetchCheName,
                String putCheName,
                String carryCheName,
                String craneCheName,
                String fromPosition,
                String toPosition
        ){
            this.visita = visita;
            this.fetchCheName = fetchCheName;
            this.putCheName = putCheName;
            this.carryCheName = carryCheName;
            this.craneCheName = craneCheName;
            this.fromPosition = fromPosition;
            this.toPosition = toPosition;
        }

        public String getFetchCheName() {
            return fetchCheName;
        }

        public String getPutCheName() {
            return putCheName;
        }

        public String getCarryCheName() {
            return carryCheName;
        }

        public String getCraneCheName() {
            return craneCheName;
        }

        public String getVisita() {
            return visita;
        }

        public String getFromPosition() {
            return fromPosition;
        }

        public String getToPosition() {
            return toPosition;
        }

        @Override
        public String toString(){
            return String.format(
              "[Visita=%s | FetchCheName=%s | CarryCheName=%s | PutCheName=%s | CraneCheName=%s | FromPosition=%s | ToPosition=%s]",
                    visita, fetchCheName, carryCheName, putCheName, craneCheName, fromPosition, toPosition
            );
        }
    }
}
