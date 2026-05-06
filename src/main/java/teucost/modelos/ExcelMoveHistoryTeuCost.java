package teucost.modelos;

public class ExcelMoveHistoryTeuCost {
    private String rutaExcel;

    private final String keyColTimeCompleted = "Time Completed";
    private final String keyColCraneCheName = "Crane CHE Name";
    private final String keyColFetchCheName = "Fetch CHE Name";
    private final String keyColPutCheName = "Put CHE Name";
    private final String keyColFromPosition = "From Position";
    private final String keyColToPosition = "To Position";
    private final String keyColCarrierVisit= "Carrier Visit";

    private int colTimeCompleted;
    private int colCraneCheName, colPutCheName,
        colFetchCheName, colCarryCheName;

    private int colFromPosition,colToPosition;

    private int colCarrierVisit;

    public ExcelMoveHistoryTeuCost(String rutaExcel){this.rutaExcel = rutaExcel;}

    public String getRutaExcel() {
        return rutaExcel;
    }

    public String getKeyColTimeCompleted() {
        return keyColTimeCompleted;
    }

    public String getKeyColCraneCheName() {
        return keyColCraneCheName;
    }

    public String getKeyColFetchCheName() {
        return keyColFetchCheName;
    }

    public String getKeyColPutCheName() {
        return keyColPutCheName;
    }

    public String getKeyColFromPosition() {
        return keyColFromPosition;
    }

    public String getKeyColToPosition() {
        return keyColToPosition;
    }

    public String getKeyColCarrierVisit() {
        return keyColCarrierVisit;
    }

    public int getColTimeCompleted() {
        return colTimeCompleted;
    }

    public int getColCraneCheName() {
        return colCraneCheName;
    }

    public int getColPutCheName() {
        return colPutCheName;
    }

    public int getColFetchCheName() {
        return colFetchCheName;
    }

    public int getColCarryCheName() {
        return colCarryCheName;
    }

    public int getColFromPosition() {
        return colFromPosition;
    }

    public int getColToPosition() {
        return colToPosition;
    }

    public int getColCarrierVisit() {
        return colCarrierVisit;
    }

    public static class MoveHistoryTeuCost{
        public final String fetCheName;
        public final String putCheName;
        public final String carryCheName;
        public final String craneCheName;
        public final String fromPosition;
        public final String toPosition;
        public final String carrierVisit;
        public final String timeCompleted;

        public MoveHistoryTeuCost(
                String fetCheName,
                String putCheName,
                String carryCheName,
                String craneCheName,
                String fromPosition,
                String toPosition,
                String timeCompleted,
                String carrierVisit
        ){
            this.carrierVisit = carrierVisit;
            this.carryCheName = carryCheName;
            this.fetCheName = fetCheName;
            this.putCheName = putCheName;
            this.fromPosition = fromPosition;
            this.craneCheName = craneCheName;
            this.toPosition = toPosition;
            this.timeCompleted = timeCompleted;
        }

        public String getFetCheName() {
            return fetCheName;
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

        public String getFromPosition() {
            return fromPosition;
        }

        public String getToPosition() {
            return toPosition;
        }

        public String getCarrierVisit() {
            return carrierVisit;
        }

        public String getTimeCompleted() {
            return timeCompleted;
        }

        @Override
        public String toString(){
            return String.format(
                    "[TimeCompleted=%s | Visita=%s | FetchCheName=%s | CarryCheName=%s | PutCheName=%s | CraneCheName=%s | FromPosition=%s | ToPosition=%s]",
                            timeCompleted, carrierVisit, fetCheName, carryCheName, putCheName, craneCheName, fromPosition, toPosition
                    );
        }
    }

}
