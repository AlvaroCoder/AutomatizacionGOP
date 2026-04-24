package postmortem.controlador;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import postmortem.modelo.ExcelMoveHistoryPostMortem;
import postmortem.modelo.ExcelMoveHistoryPostMortem.DatosMoveHistory;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class LecturaMoveHistoryPostMortem {

    private final ExcelMoveHistoryPostMortem excelMoveHistory;

    public LecturaMoveHistoryPostMortem(String rutaExcel) {
        this.excelMoveHistory = new ExcelMoveHistoryPostMortem(rutaExcel);
    }

    public List<DatosMoveHistory> extraerDatosMoveHistory() {
        List<DatosMoveHistory> resultado = new ArrayList<>();
        try {
            extraerColumnas();

            XSSFWorkbook wb = openWorkBook(excelMoveHistory.getRutaExcel());
            XSSFSheet hoja = wb.getSheetAt(0);

            int ultimaFila = hoja.getLastRowNum();

            for (int i = 1; i <= ultimaFila; i++) {
                XSSFRow row = hoja.getRow(i);
                if (row == null) continue;

                String visita = getCellString(row.getCell(excelMoveHistory.getColVisit()));
                String fetch = getCellString(row.getCell(excelMoveHistory.getColFetchCheName()));
                String put = getCellString(row.getCell(excelMoveHistory.getColPutCheName()));
                String carry = getCellString(row.getCell(excelMoveHistory.getColCarryCheName()));
                String crane = getCellString(row.getCell(excelMoveHistory.getColCraneCheName()));
                String fromPos = getCellString(row.getCell(excelMoveHistory.getColFromPosition()));
                String toPos = getCellString(row.getCell(excelMoveHistory.getColToPosition()));

                if (visita.isEmpty() && fetch.isEmpty() && carry.isEmpty()) continue;

                resultado.add(new DatosMoveHistory(visita, fetch, put, carry, crane, fromPos, toPos));
            }

            wb.close();
            System.out.println("[OK] MoveHistory extraídos: " + resultado.size());

        } catch (IOException e) {
            throw new RuntimeException("Error leyendo el Excel de MoveHistory: " + e.getMessage(), e);
        }
        return resultado;
    }

    private void extraerColumnas() throws IOException {
        XSSFWorkbook wb = openWorkBook(excelMoveHistory.getRutaExcel());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow encabezado = sheet.getRow(0);

        if (encabezado == null) {
            wb.close();
            throw new RuntimeException("El archivo no tiene encabezado en la fila 1.");
        }

        Map<String, Integer> encontradas = new HashMap<>();

        for (Cell cell : encabezado) {
            if (cell.getCellType() != CellType.STRING) continue;
            String valor = cell.getStringCellValue().trim();
            switch (valor) {
                case "Carrier Visit":   encontradas.put(excelMoveHistory.getKeyColVisita(), cell.getColumnIndex()); break;
                case "Fetch CHE Name":  encontradas.put(excelMoveHistory.getKeyColFetch(), cell.getColumnIndex()); break;
                case "Put CHE Name":    encontradas.put(excelMoveHistory.getKeyColPut(), cell.getColumnIndex()); break;
                case "Carry CHE Name":  encontradas.put(excelMoveHistory.getKeyColCarry(), cell.getColumnIndex()); break;
                case "Time Completed":  encontradas.put(excelMoveHistory.getKeyColTimeCompleted(), cell.getColumnIndex()); break;
                case "From Position":   encontradas.put(excelMoveHistory.getKeyColFromPosition(), cell.getColumnIndex()); break;
                case "To Position":     encontradas.put(excelMoveHistory.getKeyColToPosition(),  cell.getColumnIndex()); break;
            }
        }

        wb.close();

        validarColumna(encontradas, excelMoveHistory.getKeyColVisita());
        validarColumna(encontradas, excelMoveHistory.getKeyColFetch());
        validarColumna(encontradas, excelMoveHistory.getKeyColPut());
        validarColumna(encontradas, excelMoveHistory.getKeyColCarry());
        validarColumna(encontradas, excelMoveHistory.getKeyColTimeCompleted());
        validarColumna(encontradas, excelMoveHistory.getKeyColFromPosition());
        validarColumna(encontradas, excelMoveHistory.getKeyColToPosition());

        excelMoveHistory.setColVisit(encontradas.get(excelMoveHistory.getKeyColVisita()));
        excelMoveHistory.setColFetchCheName(encontradas.get(excelMoveHistory.getKeyColFetch()));
        excelMoveHistory.setColPutCheName(encontradas.get(excelMoveHistory.getKeyColPut()));
        excelMoveHistory.setColCarryCheName(encontradas.get(excelMoveHistory.getKeyColCarry()));
        excelMoveHistory.setColTimeCompleted(encontradas.get(excelMoveHistory.getKeyColTimeCompleted()));
        excelMoveHistory.setColFromPosition(encontradas.get(excelMoveHistory.getKeyColFromPosition()));
        excelMoveHistory.setColToPosition(encontradas.get(excelMoveHistory.getKeyColToPosition()));

        System.out.println("[Columnas] "
                + "Visit=" + excelMoveHistory.getColVisit()
                + " | Fetch=" + excelMoveHistory.getColFetchCheName()
                + " | Put=" + excelMoveHistory.getColPutCheName()
                + " | Carry=" + excelMoveHistory.getColCarryCheName()
                + " | Time=" + excelMoveHistory.getColTimeCompleted()
                + " | From=" + excelMoveHistory.getColFromPosition()
                + " | To=" + excelMoveHistory.getColToPosition());
    }

    private void validarColumna(Map<String, Integer> map, String key) {
        if (!map.containsKey(key)) {
            throw new RuntimeException(
                    "Columna requerida no encontrada en el encabezado: '" + key + "'"
            );
        }
    }

    private XSSFWorkbook openWorkBook(String rutaExcel) throws IOException {
        try (InputStream fis = new FileInputStream(new File(rutaExcel))) {
            return new XSSFWorkbook(fis);
        }
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                return d == Math.floor(d) ? String.valueOf((long) d) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:  return cell.getRichStringCellValue().getString().trim();
                    case NUMERIC:
                        double df = cell.getNumericCellValue();
                        return df == Math.floor(df) ? String.valueOf((long) df) : String.valueOf(df);
                    default: return "";
                }
            default: return "";
        }
    }

    private String getCellFecha(Cell cell) {
        if (cell == null) return "";
        try {
            if (cell.getCellType() == CellType.NUMERIC
                    && DateUtil.isCellDateFormatted(cell)) {
                return new SimpleDateFormat("dd/MM/yyyy HH:mm")
                        .format(cell.getDateCellValue());
            }
        } catch (Exception ignored) {}
        return getCellString(cell);
    }
}