package postmortem.modelo;

import java.util.List;

public class ExcelKPIsPostMortem {
    private double gmphData, bmphData, longGangData;
    private double craneIntensityData, gruasUtilizadasData, duracionOperacionData;

    private int movDescarga, movEmbarque, movRestiba, movTransbordo;
    private int firstLineFmove, lastLineLmove;

    private List<ExcelMoveHistoryPostMortem> excelMoveHistory;
    private List<ExcelNavesPostMortem> excelNaves;

    public ExcelKPIsPostMortem(
            List<ExcelMoveHistoryPostMortem> excelMoveHistory,
            List<ExcelNavesPostMortem> excelNaves
    ){
        this.excelMoveHistory = excelMoveHistory;
        this.excelNaves = excelNaves;
    }

    public List<ExcelMoveHistoryPostMortem> getExcelMoveHistory() {
        return excelMoveHistory;
    }

    public List<ExcelNavesPostMortem> getExcelNaves() {
        return excelNaves;
    }

    public double getGmphData() {
        return gmphData;
    }

    public void setGmphData(double gmphData) {
        this.gmphData = gmphData;
    }

    public double getBmphData() {
        return bmphData;
    }

    public void setBmphData(double bmphData) {
        this.bmphData = bmphData;
    }

    public double getLongGangData() {
        return longGangData;
    }

    public void setLongGangData(double longGangData) {
        this.longGangData = longGangData;
    }

    public double getCraneIntensityData() {
        return craneIntensityData;
    }

    public void setCraneIntensityData(double craneIntensityData) {
        this.craneIntensityData = craneIntensityData;
    }

    public double getGruasUtilizadasData() {
        return gruasUtilizadasData;
    }

    public void setGruasUtilizadasData(double gruasUtilizadasData) {
        this.gruasUtilizadasData = gruasUtilizadasData;
    }

    public double getDuracionOperacionData() {
        return duracionOperacionData;
    }

    public void setDuracionOperacionData(double duracionOperacionData) {
        this.duracionOperacionData = duracionOperacionData;
    }

    public int getMovDescarga() {
        return movDescarga;
    }

    public void setMovDescarga(int movDescarga) {
        this.movDescarga = movDescarga;
    }

    public int getMovEmbarque() {
        return movEmbarque;
    }

    public void setMovEmbarque(int movEmbarque) {
        this.movEmbarque = movEmbarque;
    }

    public int getMovRestiba() {
        return movRestiba;
    }

    public void setMovRestiba(int movRestiba) {
        this.movRestiba = movRestiba;
    }

    public int getMovTransbordo() {
        return movTransbordo;
    }

    public void setMovTransbordo(int movTransbordo) {
        this.movTransbordo = movTransbordo;
    }

    public int getFirstLineFmove() {
        return firstLineFmove;
    }

    public void setFirstLineFmove(int firstLineFmove) {
        this.firstLineFmove = firstLineFmove;
    }

    public int getLastLineLmove() {
        return lastLineLmove;
    }

    public void setLastLineLmove(int lastLineLmove) {
        this.lastLineLmove = lastLineLmove;
    }
}
