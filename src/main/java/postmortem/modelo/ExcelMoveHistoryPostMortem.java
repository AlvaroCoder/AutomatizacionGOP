package postmortem.modelo;

public class ExcelMoveHistoryPostMortem {
    private String rutaExcel;
    private int colTimeCompleted;
    private int colCraneCheName, colPutCheName,
            colFetchCheName, colCarryCheName;

    private int colVisit;

    private int colFromPosition, colToPosition;
    private int colUnitCategory;

    private final String keyColVisita = "Carrier Visit";
    private final String keyColFetch = "Fetch CHE Name";
    private final String keyColPut = "Put CHE Name";
    private final String keyColCarry = "Carry CHE Name";
    private final String keyColTimeCompleted = "Time Completed";
    private final String keyColFromPosition = "From Position";
    private final String keyColToPosition = "To Position";
    private final String keyColCraneCheName = "Crane CHE Name";
    private final String keyColUnitCategory = "Unit Category";

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

    public String getKeyColCraneCheName(){return  this.keyColCraneCheName;}

    public String getKeyColUnitCategory(){return this.keyColUnitCategory;}

    public int getColUnitCategory(){return this.colUnitCategory;}

    public void setColUnitCategory(int v){this.colUnitCategory = v;}

    public static class DatosMoveHistory{
        public final String fetchCheName;
        public final String putCheName;
        public final String carryCheName;
        public final String craneCheName;
        public final String visita;
        public final String timeCompleted;
        public final String unitCategory;

        public final String fromPosition, toPosition;

        public DatosMoveHistory(
                String visita,
                String fetchCheName,
                String putCheName,
                String carryCheName,
                String craneCheName,
                String fromPosition,
                String toPosition,
                String timeCompleted,
                String unitCategory
        ){
            this.visita = visita;
            this.fetchCheName = fetchCheName;
            this.putCheName = putCheName;
            this.carryCheName = carryCheName;
            this.craneCheName = craneCheName;
            this.fromPosition = fromPosition;
            this.toPosition = toPosition;
            this.timeCompleted = timeCompleted;
            this.unitCategory = unitCategory;
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

        public String getTimeCompleted(){ return timeCompleted; }

        public String getUnitCategory(){return unitCategory;}

        @Override
        public String toString(){
            return String.format(
              "[TimeCompleted=%s | Visita=%s | FetchCheName=%s | CarryCheName=%s | PutCheName=%s | CraneCheName=%s | FromPosition=%s | ToPosition=%s  | UnitCategory=%s]",
                    timeCompleted, visita, fetchCheName, carryCheName, putCheName, craneCheName, fromPosition, toPosition, unitCategory
            );
        }
    }
}
