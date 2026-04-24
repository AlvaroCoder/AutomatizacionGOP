package nombradas;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Lee archivos de Nombradas COSMOS (estructura de tabla dinámica).
 *
 * Por cada Excel:
 *  - Busca en col B la fila con "Total general" → toma el último valor no nulo de esa fila
 *  - Busca en col B la fila con "NOMBRES" → lee la tabla de trabajadores desde ahí
 *
 * Genera un Excel resumen con:
 *  - Hoja "Resumen": Archivo | Nave | Fecha | Turno | Total General
 *  - Hoja "Detalle": todos los trabajadores de todos los archivos concatenados
 */
public class AgruparNombradas {

    private final String rutaCarpeta;
    private final String rutaSalida;

    public AgruparNombradas(String rutaCarpeta, String rutaSalida) {
        this.rutaCarpeta = rutaCarpeta;
        this.rutaSalida  = rutaSalida;
    }

    // ==========================================================
    // PUNTO DE ENTRADA
    // ==========================================================
    public void procesar() {
        try {
            Path dir = Paths.get(rutaCarpeta).toAbsolutePath().normalize();
            if (!Files.exists(dir) || !Files.isDirectory(dir)) {
                System.out.println("[ERROR] Carpeta inválida: " + dir);
                return;
            }

            // Buscar todos los .xlsx / .xls que NO sean temporales (~$)
            List<Path> excels;
            try (Stream<Path> walk = Files.walk(dir)) {
                excels = walk
                        .filter(Files::isRegularFile)
                        .filter(p -> {
                            String n = p.getFileName().toString().toUpperCase(Locale.ROOT);
                            return !n.startsWith("~$") && (n.endsWith(".XLSX") || n.endsWith(".XLS"));
                        })
                        .sorted()
                        .collect(Collectors.toList());
            }

            if (excels.isEmpty()) {
                System.out.println("[INFO] No se encontraron archivos Excel en: " + dir);
                return;
            }

            System.out.println("\n════════ PROCESANDO " + excels.size() + " ARCHIVOS ════════");

            List<FilaResumen>  resumen  = new ArrayList<>();
            List<FilaDetalle>  detalle  = new ArrayList<>();

            for (Path excel : excels) {
                System.out.println("[LEY] " + excel.getFileName());
                try {
                    procesarArchivo(excel, resumen, detalle);
                } catch (Exception e) {
                    System.out.println("[ERROR] " + excel.getFileName() + " → " + e.getMessage());
                }
            }

            generarExcelSalida(resumen, detalle);

        } catch (Exception e) {
            System.out.println("[ERROR GENERAL] " + e.getMessage());
            e.printStackTrace();
        }
    }

    // ==========================================================
    // PROCESAR UN ARCHIVO
    // ==========================================================
    private void procesarArchivo(Path archivo,
                                 List<FilaResumen> resumen,
                                 List<FilaDetalle> detalle) throws Exception {

        try (FileInputStream fis = new FileInputStream(archivo.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            String nombreArchivo = archivo.getFileName().toString();

            // ── 1) Buscar fila "Total general" en columna B (índice 1) ──
            int filaTotal    = -1;
            int totalGeneral = 0;

            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell celdaB = row.getCell(1); // columna B
                if (celdaB == null) continue;
                String valorB = getCellString(celdaB);
                if ("total general".equalsIgnoreCase(valorB.trim())) {
                    filaTotal = i;
                    // Tomar el último valor numérico no nulo de la fila
                    totalGeneral = ultimoNumericoDeFila(row);
                    System.out.println("   → Total general encontrado en fila " + (i + 1)
                            + " = " + totalGeneral);
                    break;
                }
            }

            if (filaTotal == -1) {
                System.out.println("   [WARN] No se encontró 'Total general' en col B");
            }

            // ── 2) Buscar encabezado "NOMBRES" en columna B ──
            int filaEncabezado = -1;

            for (int i = (filaTotal > -1 ? filaTotal : sheet.getFirstRowNum());
                 i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell celdaB = row.getCell(1);
                if (celdaB == null) continue;
                String valorB = getCellString(celdaB);
                if ("nombres".equalsIgnoreCase(valorB.trim())) {
                    filaEncabezado = i;
                    System.out.println("   → Tabla detalle encontrada en fila " + (i + 1));
                    break;
                }
            }

            // ── 3) Leer encabezado para saber qué columnas existen ──
            // Estructura esperada: B=NOMBRES C=APELLIDOS D=PUESTO E=NAVE F=VIAJE G=MUELLE H=FECHA I=TURNO
            String nave   = "";
            String fecha  = "";
            String turno  = "";

            List<FilaDetalle> trabajadoresArchivo = new ArrayList<>();

            if (filaEncabezado > -1) {
                for (int i = filaEncabezado + 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String nombres   = getCellString(row.getCell(1)); // B
                    String apellidos = getCellString(row.getCell(2)); // C
                    String puesto    = getCellString(row.getCell(3)); // D
                    String naveCell  = getCellString(row.getCell(4)); // E
                    String viaje     = getCellString(row.getCell(5)); // F
                    String muelle    = getCellString(row.getCell(6)); // G
                    String fechaCell = getCellFecha(row.getCell(7));  // H
                    String turnoCell = getCellString(row.getCell(8)); // I

                    // Fila vacía → parar
                    if (nombres.isEmpty() && apellidos.isEmpty() && puesto.isEmpty()) continue;

                    // Tomar nave, fecha y turno del primer trabajador con datos
                    if (nave.isEmpty() && !naveCell.isEmpty())  nave  = naveCell;
                    if (fecha.isEmpty() && !fechaCell.isEmpty()) fecha = fechaCell;
                    if (turno.isEmpty() && !turnoCell.isEmpty()) turno = turnoCell;

                    trabajadoresArchivo.add(new FilaDetalle(
                            nombreArchivo, nombres, apellidos,
                            puesto, naveCell, viaje, muelle, fechaCell, turnoCell
                    ));
                }
            }

            System.out.println("   → Trabajadores leídos: " + trabajadoresArchivo.size());

            // ── 4) Agregar al resumen ──
            resumen.add(new FilaResumen(nombreArchivo, nave, fecha, turno, totalGeneral));
            detalle.addAll(trabajadoresArchivo);
        }
    }

    // ==========================================================
    // GENERAR EXCEL DE SALIDA
    // ==========================================================
    private void generarExcelSalida(List<FilaResumen> resumen,
                                    List<FilaDetalle> detalle) throws Exception {

        Path dirSalida = Paths.get(rutaSalida);
        if (!Files.exists(dirSalida)) Files.createDirectories(dirSalida);
        Path out = dirSalida.resolve("Resumen_Nombradas_Cosmos.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {

            CellStyle estiloHead = crearEstiloEncabezado(wb);
            CellStyle estiloData = crearEstiloDato(wb);

            // ── Hoja 1: Resumen ──
            Sheet hResumen = wb.createSheet("Resumen");
            String[] colsResumen = {"Archivo", "Nave", "Fecha", "Turno", "Total General"};
            crearEncabezado(hResumen, colsResumen, estiloHead);

            int r = 1;
            for (FilaResumen fr : resumen) {
                Row row = hResumen.createRow(r++);
                setCellData(row, 0, fr.archivo,      estiloData);
                setCellData(row, 1, fr.nave,         estiloData);
                setCellData(row, 2, fr.fecha,        estiloData);
                setCellData(row, 3, fr.turno,        estiloData);
                setCellData(row, 4, fr.totalGeneral, estiloData);
            }
            for (int i = 0; i < colsResumen.length; i++) hResumen.autoSizeColumn(i);
            hResumen.createFreezePane(0, 1);

            // ── Hoja 2: Detalle ──
            Sheet hDetalle = wb.createSheet("Detalle");
            String[] colsDetalle = {"Archivo", "Nombres", "Apellidos",
                    "Puesto", "Nave", "Viaje", "Muelle", "Fecha", "Turno"};
            crearEncabezado(hDetalle, colsDetalle, estiloHead);

            int d = 1;
            for (FilaDetalle fd : detalle) {
                Row row = hDetalle.createRow(d++);
                setCellData(row, 0, fd.archivo,   estiloData);
                setCellData(row, 1, fd.nombres,   estiloData);
                setCellData(row, 2, fd.apellidos, estiloData);
                setCellData(row, 3, fd.puesto,    estiloData);
                setCellData(row, 4, fd.nave,      estiloData);
                setCellData(row, 5, fd.viaje,     estiloData);
                setCellData(row, 6, fd.muelle,    estiloData);
                setCellData(row, 7, fd.fecha,     estiloData);
                setCellData(row, 8, fd.turno,     estiloData);
            }
            for (int i = 0; i < colsDetalle.length; i++) hDetalle.autoSizeColumn(i);
            hDetalle.createFreezePane(0, 1);

            try (FileOutputStream fos = new FileOutputStream(out.toFile())) {
                wb.write(fos);
            }
        }

        System.out.println("\n════════ RESUMEN FINAL ════════");
        System.out.println("✅ Archivos procesados : " + resumen.size());
        System.out.println("✅ Trabajadores totales: " + detalle.size());
        System.out.println("💾 Guardado en         : " + out);
    }

    // ==========================================================
    // UTILIDADES DE LECTURA DE CELDAS
    // ==========================================================

    /** Devuelve el último valor numérico entero no nulo de una fila. */
    private int ultimoNumericoDeFila(Row row) {
        int ultimo = 0;
        for (Cell cell : row) {
            if (cell == null) continue;
            try {
                switch (cell.getCellType()) {
                    case NUMERIC:
                        ultimo = (int) cell.getNumericCellValue();
                        break;
                    case FORMULA:
                        if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                            ultimo = (int) cell.getNumericCellValue();
                        }
                        break;
                    default: break;
                }
            } catch (Exception ignored) {}
        }
        return ultimo;
    }

    /** Convierte una celda a String limpio. */
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

    /** Lee una celda de fecha y la formatea como dd/MM/yyyy. */
    private String getCellFecha(Cell cell) {
        if (cell == null) return "";
        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                return new SimpleDateFormat("dd/MM/yyyy").format(date);
            }
        } catch (Exception ignored) {}
        return getCellString(cell);
    }

    // ==========================================================
    // UTILIDADES DE ESCRITURA
    // ==========================================================
    private void crearEncabezado(Sheet sheet, String[] cols, CellStyle estilo) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < cols.length; i++) {
            Cell c = row.createCell(i);
            c.setCellValue(cols[i]);
            c.setCellStyle(estilo);
        }
    }

    private void setCellData(Row row, int col, String val, CellStyle estilo) {
        Cell c = row.createCell(col);
        c.setCellValue(val != null ? val : "");
        c.setCellStyle(estilo);
    }

    private void setCellData(Row row, int col, int val, CellStyle estilo) {
        Cell c = row.createCell(col);
        c.setCellValue(val);
        c.setCellStyle(estilo);
    }

    private CellStyle crearEstiloEncabezado(XSSFWorkbook wb) {
        Font f = wb.createFont();
        f.setBold(true);
        f.setColor(IndexedColors.WHITE.getIndex());
        CellStyle s = wb.createCellStyle();
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        return s;
    }

    private CellStyle crearEstiloDato(XSSFWorkbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    // ==========================================================
    // POJOs
    // ==========================================================
    private static class FilaResumen {
        String archivo, nave, fecha, turno;
        int totalGeneral;
        FilaResumen(String archivo, String nave, String fecha, String turno, int totalGeneral) {
            this.archivo = archivo; this.nave = nave; this.fecha = fecha;
            this.turno = turno; this.totalGeneral = totalGeneral;
        }
    }

    private static class FilaDetalle {
        String archivo, nombres, apellidos, puesto, nave, viaje, muelle, fecha, turno;
        FilaDetalle(String archivo, String nombres, String apellidos,
                    String puesto, String nave, String viaje,
                    String muelle, String fecha, String turno) {
            this.archivo = archivo; this.nombres = nombres; this.apellidos = apellidos;
            this.puesto = puesto; this.nave = nave; this.viaje = viaje;
            this.muelle = muelle; this.fecha = fecha; this.turno = turno;
        }
    }
}