package postmortem.controlador;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import postmortem.modelo.ExcelNavesPostMortem;
import postmortem.modelo.ExcelNavesPostMortem.DatosNave;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class LecturaNaves {

    private final ExcelNavesPostMortem excelNaves;

    public LecturaNaves(String rutaNaves) {
        this.excelNaves = new ExcelNavesPostMortem(rutaNaves);
    }

    public List<DatosNave> extraerDatosNave() {
        List<DatosNave> resultado = new ArrayList<>();
        try {
            extraerColumnas();

            XSSFWorkbook wb = openWorkBook(excelNaves.getRutaExcel());
            XSSFSheet hoja  = wb.getSheetAt(0);

            int primeraFilaDatos = 1;
            int ultimaFila       = hoja.getLastRowNum();

            for (int i = primeraFilaDatos; i <= ultimaFila; i++) {
                XSSFRow row = hoja.getRow(i);
                if (row == null) continue;

                String visita = getCellString(row.getCell(excelNaves.getColVisita()));
                String linea = getCellString(row.getCell(excelNaves.getColLinea()));
                String servicio = getCellString(row.getCell(excelNaves.getColServicio()));
                String inicio = getCellFecha (row.getCell(excelNaves.getColInicioOperaciones()));
                String fin = getCellFecha (row.getCell(excelNaves.getColFinOperaciones()));
                String berthATA = getCellFecha(row.getCell(excelNaves.getColBerthATA()));
                String berthATD = getCellFecha(row.getCell(excelNaves.getColBerthATD()));

                if (visita.isEmpty() && linea.isEmpty() && servicio.isEmpty()) continue;

                resultado.add(new DatosNave(visita, linea, servicio, inicio, fin, berthATA, berthATD));
            }

            wb.close();

            System.out.println("[OK] Naves extraídas: " + resultado.size());

        } catch (IOException e) {
            throw new RuntimeException("Error leyendo el Excel de naves: " + e.getMessage(), e);
        }
        return resultado;
    }

    private void extraerColumnas() throws IOException {
        XSSFWorkbook wb    = openWorkBook(excelNaves.getRutaExcel());
        XSSFSheet sheet    = wb.getSheetAt(0);
        XSSFRow   encabezado = sheet.getRow(0);

        if (encabezado == null) {
            wb.close();
            throw new RuntimeException("El archivo no tiene fila de encabezado en la fila 1.");
        }

        Map<String, Integer> encontradas = new HashMap<>();

        for (Cell cell : encabezado) {
            if (cell.getCellType() != CellType.STRING) continue;
            String valor = cell.getStringCellValue().trim();
            switch (valor) {
                case "Start Work":
                    encontradas.put(excelNaves.getKeyColInicio(), cell.getColumnIndex());
                    break;
                case "End Work":
                    encontradas.put(excelNaves.getKeyColFin(), cell.getColumnIndex());
                    break;
                case "Visit":
                    encontradas.put(excelNaves.getKeyColVisita(), cell.getColumnIndex());
                    break;
                case "Line":
                    encontradas.put(excelNaves.getKeyColLinea(), cell.getColumnIndex());
                    break;
                case "Service":
                    encontradas.put(excelNaves.getKeyColServicio(), cell.getColumnIndex());
                    break;
                case  "ATA":
                    encontradas.put(excelNaves.getKeyColBerthATA(), cell.getColumnIndex());
                    break;
                case "ATD":
                    encontradas.put(excelNaves.getKeyColBerthATD(), cell.getColumnIndex());
                    break;
            }
        }

        wb.close();

        validarColumna(encontradas, excelNaves.getKeyColVisita());
        validarColumna(encontradas, excelNaves.getKeyColLinea());
        validarColumna(encontradas, excelNaves.getKeyColServicio());
        validarColumna(encontradas, excelNaves.getKeyColInicio());
        validarColumna(encontradas, excelNaves.getKeyColFin());
        validarColumna(encontradas, excelNaves.getKeyColBerthATA());
        validarColumna(encontradas, excelNaves.getKeyColBerthATD());

        excelNaves.setColVisita(encontradas.get(excelNaves.getKeyColVisita()));
        excelNaves.setColLinea(encontradas.get(excelNaves.getKeyColLinea()));
        excelNaves.setColServicio(encontradas.get(excelNaves.getKeyColServicio()));
        excelNaves.setColInicioOperaciones(encontradas.get(excelNaves.getKeyColInicio()));
        excelNaves.setColFinOperaciones(encontradas.get(excelNaves.getKeyColFin()));
        excelNaves.setColBerthATA(encontradas.get(excelNaves.getKeyColBerthATA()));
        excelNaves.setColBerthATD(encontradas.get(excelNaves.getKeyColBerthATD()));

        System.out.println("[Columnas] Visit=" + excelNaves.getColVisita()
                + " | Line="+ excelNaves.getColLinea()
                + " | Service="+ excelNaves.getColServicio()
                + " | Start="+ excelNaves.getColInicioOperaciones()
                + " | End="+ excelNaves.getColFinOperaciones());
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

    /** Lee cualquier celda como String limpio. */
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

    /**
     * Lee una celda de fecha.
     * Si POI la reconoce como fecha → formatea dd/MM/yyyy HH:mm
     * Si es texto → lo devuelve tal cual
     */
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