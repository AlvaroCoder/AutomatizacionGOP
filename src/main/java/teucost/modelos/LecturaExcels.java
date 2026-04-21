package teucost.modelos;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.*;

public class LecturaExcels {

    private String rutaExcel;
    private Workbook workbook;


    public LecturaExcels(String rutaExcel) throws IOException {
        this.rutaExcel = rutaExcel;
        this.workbook  = abrirWorkbook(rutaExcel);
    }

    private Workbook abrirWorkbook(String ruta) throws IOException {
        ZipSecureFile.setMaxFileCount(5000);
        try (FileInputStream fis = new FileInputStream(new File(ruta))) {
            String rutaLower = ruta.toLowerCase();
            if (rutaLower.endsWith(".xlsx") || rutaLower.endsWith(".xlsm")) {
                return new XSSFWorkbook(fis);
            } else if (rutaLower.endsWith(".xls")) {
                return new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException(
                        "Formato no soportado. Use .xlsx, .xlsm o .xls");
            }
        }
    }

    public Row insertarFilaEnTabla(int indiceHoja, int indiceTabla) {
        if (!(workbook instanceof XSSFWorkbook)) {
            throw new IllegalStateException(
                    "insertarFilaEnTabla solo soporta archivos .xlsx (XSSFWorkbook)");
        }

        XSSFWorkbook xssfWb = (XSSFWorkbook) workbook;
        XSSFSheet    hoja   = xssfWb.getSheetAt(indiceHoja);
        XSSFTable    tabla  = hoja.getTables().get(indiceTabla);

        AreaReference areaActual = new AreaReference(
                tabla.getCTTable().getRef(), xssfWb.getSpreadsheetVersion());

        CellReference esquinaInicio = areaActual.getFirstCell();
        CellReference esquinaFin    = areaActual.getLastCell();

        int filaFinTabla = esquinaFin.getRow();
        int colInicio    = esquinaInicio.getCol();
        int colFin       = esquinaFin.getCol();
        int nuevaFilaIdx = filaFinTabla + 1;

        Row filaAnterior = hoja.getRow(filaFinTabla);
        Row nuevaFila    = hoja.createRow(nuevaFilaIdx);

        if (filaAnterior != null) {
            nuevaFila.setHeight(filaAnterior.getHeight());
            for (int c = colInicio; c <= colFin; c++) {
                Cell celdaOrigen  = filaAnterior.getCell(c);
                Cell celdaDestino = nuevaFila.createCell(c);
                if (celdaOrigen != null && celdaOrigen.getCellStyle() != null) {
                    XSSFCellStyle estiloCopiado = xssfWb.createCellStyle();
                    estiloCopiado.cloneStyleFrom(celdaOrigen.getCellStyle());
                    celdaDestino.setCellStyle(estiloCopiado);
                }
            }
        }

        CellReference nuevaEsquinaFin = new CellReference(nuevaFilaIdx, colFin);
        AreaReference nuevaArea = new AreaReference(
                esquinaInicio, nuevaEsquinaFin, xssfWb.getSpreadsheetVersion());
        tabla.getCTTable().setRef(nuevaArea.formatAsString());

        if (tabla.getCTTable().getAutoFilter() != null) {
            tabla.getCTTable().getAutoFilter().setRef(nuevaArea.formatAsString());
        }

        System.out.println("[INFO] Tabla expandida: " + areaActual.formatAsString()
                + " → " + nuevaArea.formatAsString()
                + " | Nueva fila idx: " + nuevaFilaIdx);

        return nuevaFila;
    }

    public Row insertarFilaEnTabla(int indiceHoja) {
        return insertarFilaEnTabla(indiceHoja, 0);
    }

    public Row agregarFilaSinTabla(int indiceHoja, int colReferencia) {
        Sheet hoja = workbook.getSheetAt(indiceHoja);
        if (hoja == null) {
            hoja = workbook.createSheet();
        }

        int ultimaFila = -1;
        for (int f = hoja.getLastRowNum(); f >= 0; f--) {
            Row row = hoja.getRow(f);
            if (row == null) continue;
            Cell celda = row.getCell(colReferencia);
            if (celda != null
                    && celda.getCellType() != CellType.BLANK
                    && !celda.toString().trim().isEmpty()) {
                ultimaFila = f;
                break;
            }
        }

        int nuevaFilaIdx = ultimaFila + 1;
        Row nuevaFila = hoja.getRow(nuevaFilaIdx);
        if (nuevaFila == null) {
            nuevaFila = hoja.createRow(nuevaFilaIdx);
        }

        System.out.println("[INFO] agregarFilaSinTabla → nueva fila idx: " + nuevaFilaIdx);
        return nuevaFila;
    }

    public boolean hojaContieneEncabezados(int indiceHoja, int colReferencia) {
        try {
            if (indiceHoja >= workbook.getNumberOfSheets()) return false;
            Sheet hoja = workbook.getSheetAt(indiceHoja);
            if (hoja == null || hoja.getPhysicalNumberOfRows() == 0) return false;
            Row primera = hoja.getRow(0);
            if (primera == null) return false;
            Cell celda = primera.getCell(colReferencia);
            return celda != null
                    && celda.getCellType() != CellType.BLANK
                    && !celda.toString().trim().isEmpty();
        } catch (Exception e) {
            return false;
        }
    }

    public boolean hojaExiste(int indiceHoja) {
        return indiceHoja >= 0 && indiceHoja < workbook.getNumberOfSheets();
    }

    public Sheet garantizarHoja(int indiceHoja, String nombre) {
        while (workbook.getNumberOfSheets() <= indiceHoja) {
            workbook.createSheet();
        }
        Sheet hoja = workbook.getSheetAt(indiceHoja);
        // Renombrar solo si el nombre actual es genérico (Sheet0, Sheet1…)
        String nombreActual = workbook.getSheetName(indiceHoja);
        if (nombreActual.matches("Sheet\\d+")) {
            workbook.setSheetName(indiceHoja, nombre);
        }
        return hoja;
    }

    public int obtenerPrimeraFilaVaciaEnTabla(int indiceHoja, int indiceTabla,
                                              int colReferencia) {
        if (!(workbook instanceof XSSFWorkbook)) return -1;

        XSSFWorkbook xssfWb = (XSSFWorkbook) workbook;
        XSSFSheet    hoja   = xssfWb.getSheetAt(indiceHoja);
        XSSFTable    tabla  = hoja.getTables().get(indiceTabla);

        AreaReference area = new AreaReference(
                tabla.getCTTable().getRef(), xssfWb.getSpreadsheetVersion());

        int filaInicioDatos = area.getFirstCell().getRow() + 1;
        int filaFin         = area.getLastCell().getRow();

        for (int f = filaInicioDatos; f <= filaFin; f++) {
            Row row = hoja.getRow(f);
            if (row == null) return f;
            Cell celda = row.getCell(colReferencia);
            if (celda == null
                    || celda.getCellType() == CellType.BLANK
                    || celda.toString().trim().isEmpty()) {
                return f;
            }
        }
        return filaFin + 1;
    }

    public Row obtenerUltimaFilaConDatosEnTabla(int indiceHoja, int indiceTabla,
                                                int colReferencia) {
        if (!(workbook instanceof XSSFWorkbook)) return null;

        XSSFWorkbook xssfWb = (XSSFWorkbook) workbook;
        XSSFSheet    hoja   = xssfWb.getSheetAt(indiceHoja);
        XSSFTable    tabla  = hoja.getTables().get(indiceTabla);

        AreaReference area = new AreaReference(
                tabla.getCTTable().getRef(), xssfWb.getSpreadsheetVersion());

        int filaInicioDatos = area.getFirstCell().getRow() + 1;
        int filaFin         = area.getLastCell().getRow();

        Row ultimaConDatos = null;
        for (int f = filaInicioDatos; f <= filaFin; f++) {
            Row row = hoja.getRow(f);
            if (row == null) continue;
            Cell celda = row.getCell(colReferencia);
            if (celda != null
                    && celda.getCellType() != CellType.BLANK
                    && !celda.toString().trim().isEmpty()) {
                ultimaConDatos = row;
            }
        }
        return ultimaConDatos;
    }

    public String leerCelda(String nombreHoja, int fila, int columna) {
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return "";
        Row row = hoja.getRow(fila);
        if (row == null) return "";
        return obtenerValorCelda(row.getCell(columna));
    }

    public List<String> leerFila(String nombreHoja, int fila) {
        List<String> valores = new ArrayList<>();
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return valores;
        Row row = hoja.getRow(fila);
        if (row == null) return valores;
        for (int i = 0; i < row.getLastCellNum(); i++) {  // lastCellNum es exclusivo: correcto
            valores.add(obtenerValorCelda(row.getCell(i)));
        }
        return valores;
    }

    public List<String> leerColumna(String nombreHoja, int columna) {
        List<String> valores = new ArrayList<>();
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return valores;
        for (Row row : hoja) {
            valores.add(obtenerValorCelda(row.getCell(columna)));
        }
        return valores;
    }

    public Map<String, String> leerRango(String nombreHoja,
                                         int filaInicio, int filaFin,
                                         int colInicio, int colFin) {
        Map<String, String> datos = new LinkedHashMap<>();
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return datos;
        for (int f = filaInicio; f <= filaFin; f++) {
            Row row = hoja.getRow(f);
            for (int c = colInicio; c <= colFin; c++) {
                String valor = (row != null) ? obtenerValorCelda(row.getCell(c)) : "";
                datos.put(f + "," + c, valor);
            }
        }
        return datos;
    }

    public List<List<String>> leerHojaCompleta(String nombreHoja) {
        List<List<String>> tabla = new ArrayList<>();
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return tabla;
        for (Row row : hoja) {
            List<String> fila = new ArrayList<>();
            int lastCol = row.getLastCellNum();   // exclusivo — no sumar 1
            for (int c = 0; c < lastCol; c++) {
                fila.add(obtenerValorCelda(row.getCell(c)));
            }
            tabla.add(fila);
        }
        return tabla;
    }

    public String leerCeldaPorIndice(int indiceHoja, int fila, int columna) {
        Sheet hoja = workbook.getSheetAt(indiceHoja);
        if (hoja == null) return "";
        Row row = hoja.getRow(fila);
        if (row == null) return "";
        return obtenerValorCelda(row.getCell(columna));
    }

    public List<String> leerFilaEnRango(String nombreHoja, int fila,
                                        int colInicio, int colFin) {
        List<String> valores = new ArrayList<>();
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return valores;
        Row row = hoja.getRow(fila);
        if (row == null) {
            for (int c = colInicio; c <= colFin; c++) valores.add("");
            return valores;
        }
        for (int c = colInicio; c <= colFin; c++) {
            valores.add(obtenerValorCelda(row.getCell(c)));
        }
        return valores;
    }

    public void escribirCelda(String nombreHoja, int fila, int columna, String valor) {
        Sheet hoja = obtenerOCrearHoja(nombreHoja);
        obtenerOCrearCelda(obtenerOCrearFila(hoja, fila), columna).setCellValue(valor);
    }

    public void escribirCeldaNumero(String nombreHoja, int fila, int columna, double valor) {
        Sheet hoja = obtenerOCrearHoja(nombreHoja);
        obtenerOCrearCelda(obtenerOCrearFila(hoja, fila), columna).setCellValue(valor);
    }

    public void escribirFila(String nombreHoja, int fila, int colInicio,
                             List<String> valores) {
        Sheet hoja = obtenerOCrearHoja(nombreHoja);
        Row row = obtenerOCrearFila(hoja, fila);
        for (int i = 0; i < valores.size(); i++) {
            obtenerOCrearCelda(row, colInicio + i).setCellValue(valores.get(i));
        }
    }

    public void pegarTabla(String nombreHoja, int filaInicio, int colInicio,
                           List<List<String>> datos) {
        Sheet hoja = obtenerOCrearHoja(nombreHoja);
        for (int f = 0; f < datos.size(); f++) {
            Row row = obtenerOCrearFila(hoja, filaInicio + f);
            List<String> filaData = datos.get(f);
            for (int c = 0; c < filaData.size(); c++) {
                obtenerOCrearCelda(row, colInicio + c).setCellValue(filaData.get(c));
            }
        }
    }

    public void copiarRangoDentroDelArchivo(String hojaOrigen, String hojaDestino,
                                            int filaOrigenIni, int filaOrigenFin,
                                            int colOrigenIni, int colOrigenFin,
                                            int filaDestIni, int colDestIni) {
        Map<String, String> datos = leerRango(hojaOrigen,
                filaOrigenIni, filaOrigenFin, colOrigenIni, colOrigenFin);
        Sheet destino = obtenerOCrearHoja(hojaDestino);
        for (Map.Entry<String, String> entry : datos.entrySet()) {
            String[] coords = entry.getKey().split(",");
            int f = Integer.parseInt(coords[0]) - filaOrigenIni + filaDestIni;
            int c = Integer.parseInt(coords[1]) - colOrigenIni + colDestIni;
            obtenerOCrearCelda(obtenerOCrearFila(destino, f), c)
                    .setCellValue(entry.getValue());
        }
    }

    public void copiarRangoAOtroArchivo(String hojaOrigen, String rutaDestino,
                                        String hojaDestino,
                                        int filaOrigenIni, int filaOrigenFin,
                                        int colOrigenIni, int colOrigenFin,
                                        int filaDestIni, int colDestIni) throws IOException {
        Map<String, String> datos = leerRango(hojaOrigen,
                filaOrigenIni, filaOrigenFin, colOrigenIni, colOrigenFin);

        File archivoDestino = new File(rutaDestino);
        Workbook wbDest;
        if (archivoDestino.exists()) {
            try (FileInputStream fis = new FileInputStream(archivoDestino)) {
                wbDest = (rutaDestino.endsWith(".xlsx") || rutaDestino.endsWith(".xlsm"))
                        ? new XSSFWorkbook(fis)
                        : new HSSFWorkbook(fis);
            }
        } else {
            wbDest = (rutaDestino.endsWith(".xlsx") || rutaDestino.endsWith(".xlsm"))
                    ? new XSSFWorkbook()
                    : new HSSFWorkbook();
        }

        Sheet hojaDest = wbDest.getSheet(hojaDestino);
        if (hojaDest == null) hojaDest = wbDest.createSheet(hojaDestino);

        for (Map.Entry<String, String> entry : datos.entrySet()) {
            String[] coords = entry.getKey().split(",");
            int f = Integer.parseInt(coords[0]) - filaOrigenIni + filaDestIni;
            int c = Integer.parseInt(coords[1]) - colOrigenIni + colDestIni;
            Row row  = hojaDest.getRow(f);
            if (row == null) row = hojaDest.createRow(f);
            Cell cell = row.getCell(c);
            if (cell == null) cell = row.createCell(c);
            cell.setCellValue(entry.getValue());
        }

        try (FileOutputStream fos = new FileOutputStream(archivoDestino)) {
            wbDest.write(fos);
        }
        wbDest.close();
    }

    public void guardar() throws IOException {
        guardarComo(this.rutaExcel);
    }

    public void guardarComo(String rutaNueva) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(new File(rutaNueva))) {
            workbook.write(fos);
        }
    }

    public List<String> obtenerNombresDeHojas() {
        List<String> nombres = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            nombres.add(workbook.getSheetName(i));
        }
        return nombres;
    }

    public int obtenerTotalFilas(String nombreHoja) {
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null || hoja.getPhysicalNumberOfRows() == 0) return 0;
        return hoja.getLastRowNum() + 1;
    }

    public int obtenerTotalColumnas(String nombreHoja) {
        Sheet hoja = workbook.getSheet(nombreHoja);
        if (hoja == null) return 0;
        Row primeraFila = hoja.getRow(0);
        return (primeraFila == null) ? 0 : primeraFila.getLastCellNum();
    }

    public void cerrar() throws IOException {
        if (workbook != null) workbook.close();
    }

    private String obtenerValorCelda(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double num = cell.getNumericCellValue();
                if (num == Math.floor(num) && !Double.isInfinite(num)) {
                    return String.valueOf((long) num);
                }
                return String.valueOf(num);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC:
                        // CORRECCIÓN: verificar fecha antes de leer como fecha
                        if (DateUtil.isCellDateFormatted(cell)) {
                            return cell.getDateCellValue().toString();
                        }
                        double numFormula = cell.getNumericCellValue();
                        if (numFormula == Math.floor(numFormula)
                                && !Double.isInfinite(numFormula)) {
                            return String.valueOf((long) numFormula);
                        }
                        return String.valueOf(numFormula);
                    case STRING:
                        return cell.getRichStringCellValue().getString();
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                    case ERROR:
                        return "";
                    default:
                        return "";
                }
            case BLANK:
            default:
                return "";
        }
    }

    private Sheet obtenerOCrearHoja(String nombre) {
        Sheet hoja = workbook.getSheet(nombre);
        return (hoja != null) ? hoja : workbook.createSheet(nombre);
    }

    private Row obtenerOCrearFila(Sheet hoja, int indice) {
        Row row = hoja.getRow(indice);
        return (row != null) ? row : hoja.createRow(indice);
    }

    private Cell obtenerOCrearCelda(Row row, int indice) {
        Cell cell = row.getCell(indice);
        return (cell != null) ? cell : row.createCell(indice);
    }

    // =========================================================
    //  GETTERS
    // =========================================================

    public String getRutaExcel()   { return rutaExcel; }
    public Workbook getWorkbook()  { return workbook;  }
}
