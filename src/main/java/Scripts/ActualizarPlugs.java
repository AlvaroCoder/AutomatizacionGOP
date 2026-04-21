package Scripts;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Locale;

public class ActualizarPlugs {
    private String CarpetaReportes;
    private String RutaPlugsExcel;
    private HashMap<String, Double> objetosPlugs;

    private int mesReferencia;
    private int anioReferencia;

    public ActualizarPlugs(String RutaCarpetaReportes, String RutaPlugsExcel){
        this.CarpetaReportes = RutaCarpetaReportes;
        this.RutaPlugsExcel = RutaPlugsExcel;

        // Inicialización por defecto con fecha actual
        LocalDate hoy = LocalDate.now();
        this.mesReferencia = hoy.getMonthValue();
        this.anioReferencia = hoy.getYear();
    }

    public void setPeriodo(int mes, int anio) {
        this.mesReferencia = mes;
        this.anioReferencia = anio;
    }

    private static ArrayList<String> extraerTomasBloques(){
        try(
                FileInputStream tomasReeferStream = new FileInputStream("Q:\\PRACTICANTE PLANEAMIENTO\\RENATO\\DASHBOARD OK\\TOMAS REEFER.xlsx");
                XSSFWorkbook tomasReeferWorkbook = new XSSFWorkbook(tomasReeferStream)
                ) {
            ArrayList <String> listaBloques = new ArrayList<>();
            XSSFSheet tomasReeferSheet = tomasReeferWorkbook.getSheetAt(0);
            XSSFTable tablaTomasReefer = tomasReeferSheet.getTables().get(0);

            int inicio = tablaTomasReefer.getStartRowIndex();
            int fin = tablaTomasReefer.getEndRowIndex();

            for (int i = inicio +1; i <= fin; i++) {
                Cell celdaBloque = tomasReeferSheet.getRow(i).getCell(1);
                String nombreBloque = celdaBloque.getStringCellValue();
                if (!nombreBloque.equals("BLOCK")){
                    listaBloques.add(nombreBloque);
                }
            }

            return listaBloques;
        }
        catch (Exception e ){
            System.out.println("e = " + e.getMessage());
        }
        return null;
    }

    private void actualizarTabla(int fechaDia, XSSFWorkbook plugsExcelWorkbook) {
        XSSFSheet plugsExcelSheet = plugsExcelWorkbook.getSheetAt(0);

        // 1. Validar directorio de reportes
        Path directorioReportes = Paths.get(this.CarpetaReportes);
        if (!Files.exists(directorioReportes) || !Files.isDirectory(directorioReportes)) {
            throw new IllegalArgumentException("La ruta no es válida: " + CarpetaReportes);
        }
        File folderReportes = new File(this.CarpetaReportes);
        File[] listOfFiles = folderReportes.listFiles();

        ArrayList<String> listaNombreBloques = extraerTomasBloques();

        // Reiniciamos el mapa para el día actual del bucle
        this.objetosPlugs = new HashMap<>();

        // 2. Buscar el archivo .xls que corresponde al día solicitado
        if (listOfFiles != null) {
            for (File archivoReporte : listOfFiles) {
                String nombre = archivoReporte.getName();
                if (archivoReporte.isFile() && nombre.endsWith(".xls")) {
                    try {
                        int diaArchivo = Integer.parseInt(nombre.substring(0, 2));
                        if (fechaDia == diaArchivo) {
                            try (FileInputStream fis = new FileInputStream(archivoReporte);
                                 HSSFWorkbook archivoWorkbook = new HSSFWorkbook(fis)) {

                                HSSFSheet archivoSheet = archivoWorkbook.getSheetAt(0);
                                int filaInicioTabla = -1;

                                // Buscar la fila de inicio
                                for (Row fila : archivoSheet) {
                                    Cell celda0 = fila.getCell(0);
                                    if (celda0 != null && celda0.getCellType() == CellType.STRING &&
                                            celda0.getStringCellValue().contains("Summary for Reefers in Yard by Block")) {
                                        filaInicioTabla = fila.getRowNum();
                                        break;
                                    }
                                }

                                if (filaInicioTabla != -1) {
                                    int numHeader = filaInicioTabla + 1;
                                    int colTotal = 0;
                                    // Buscar columna "Total"
                                    for (Cell col : archivoSheet.getRow(numHeader)) {
                                        if (col.getCellType() == CellType.STRING && col.getStringCellValue().equals("Total")) {
                                            colTotal = col.getColumnIndex();
                                            break;
                                        }
                                    }

                                    // Extraer datos de los bloques
                                    int filaAct = numHeader + 1;
                                    while (true) {
                                        Row fila = archivoSheet.getRow(filaAct);
                                        if (fila == null) break;
                                        Cell cBloque = fila.getCell(0);
                                        if (cBloque == null) break;

                                        String nomBloque = cBloque.getStringCellValue().trim();
                                        if (listaNombreBloques != null && listaNombreBloques.contains(nomBloque)) {
                                            this.objetosPlugs.put(nomBloque, fila.getCell(colTotal).getNumericCellValue());
                                        }
                                        if (nomBloque.equalsIgnoreCase("Total")) break;
                                        filaAct++;
                                    }
                                }
                            } catch (Exception e) {
                                System.out.println("Error procesando archivo " + nombre + ": " + e.getMessage());
                            }
                        }
                    } catch (NumberFormatException ignored) {}
                }
            }
        }

        // 3. Modificar la Tabla en el Excel Destino (Plugs)
        XSSFTable tablePlugs = plugsExcelSheet.getTables().get(0);
        int startRow = tablePlugs.getStartRowIndex();
        int endRow = tablePlugs.getEndRowIndex();
        int startCol = tablePlugs.getStartColIndex();
        int endCol = tablePlugs.getEndColIndex();
        int nuevaCol = endCol + 1;

        // Expandir el área de la tabla
        tablePlugs.setArea(new AreaReference(
                new CellReference(startRow, startCol),
                new CellReference(endRow, nuevaCol),
                SpreadsheetVersion.EXCEL2007
        ));

        // 4. Configurar Fechas Dinámicas (USANDO EL PERIODO DE REFERENCIA)
        LocalDate fechaRef = LocalDate.of(this.anioReferencia, this.mesReferencia, fechaDia);

        // Header 1: "15 enero"
        String txtFecha = fechaRef.getDayOfMonth() + " " +
                fechaRef.getMonth().getDisplayName(TextStyle.FULL, new Locale("es", "ES"));
        Row rHeader = plugsExcelSheet.getRow(startRow);
        Cell cHeader = rHeader.createCell(nuevaCol);
        cHeader.setCellStyle(rHeader.getCell(endCol).getCellStyle());
        cHeader.setCellValue(txtFecha);

        // Header 2: "Monday"
        String txtDia = fechaRef.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH);
        Row rDia = plugsExcelSheet.getRow(startRow + 1);
        Cell cDia = rDia.createCell(nuevaCol);
        cDia.setCellStyle(rDia.getCell(endCol).getCellStyle());
        cDia.setCellValue(txtDia);

        // 5. Llenado de Datos y Fórmulas
        for (int i = startRow + 2; i <= endRow; i++) {
            Row fila = plugsExcelSheet.getRow(i);
            Cell nuevaCelda = fila.createCell(nuevaCol);
            CellStyle estiloModelo = fila.getCell(endCol).getCellStyle();
            nuevaCelda.setCellStyle(estiloModelo);

            String nombreBucket = fila.getCell(2) != null ? fila.getCell(2).getStringCellValue() : "";

            if (nombreBucket.isEmpty()) {
                // Lógica de desplazamiento de fórmulas
                Cell celdaModelo = fila.getCell(endCol);
                if (celdaModelo != null && celdaModelo.getCellType() == CellType.FORMULA) {
                    int sIdx = plugsExcelWorkbook.getSheetIndex(plugsExcelSheet);
                    XSSFEvaluationWorkbook evalWB = XSSFEvaluationWorkbook.create(plugsExcelWorkbook);
                    Ptg[] ptgs = FormulaParser.parse(celdaModelo.getCellFormula(), evalWB, FormulaType.CELL, sIdx);

                    FormulaShifter shifter = FormulaShifter.createForColumnShift(sIdx, null, endCol, endCol, 1, SpreadsheetVersion.EXCEL2007);
                    shifter.adjustFormula(ptgs, sIdx);

                    nuevaCelda.setCellFormula(FormulaRenderer.toFormulaString(evalWB, ptgs));
                }
            } else {
                // Llenar con datos del HashMap o 0 si no existe
                if (this.objetosPlugs.containsKey(nombreBucket)) {
                    nuevaCelda.setCellValue(this.objetosPlugs.get(nombreBucket));
                } else {
                    nuevaCelda.setCellValue(0);
                }
            }
        }
        System.out.println("Día " + fechaDia + " de " + txtFecha + " procesado correctamente.");
    }

    public void actualizarTablaDiaria() {
        try(
                FileInputStream plugsExcelStream = new FileInputStream(this.RutaPlugsExcel);
                XSSFWorkbook plugsExcelWorkbook = new XSSFWorkbook(plugsExcelStream);
                ) {

            int fechaHoy = LocalDate.now().getDayOfMonth();
            this.actualizarTabla(fechaHoy, plugsExcelWorkbook);

            try (FileOutputStream plugsOutput = new FileOutputStream(this.RutaPlugsExcel)) {
                plugsExcelWorkbook.write(plugsOutput);
            } catch (FileNotFoundException ex) {
                throw new RuntimeException(ex);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            System.out.println("Se cerro el Excel !!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void actualizarTablaPorRangoFechas(int diaInicio, int diaFin){
        try(
                FileInputStream plugsExcelStream = new FileInputStream(this.RutaPlugsExcel);
                XSSFWorkbook plugsExcelWorkbook = new XSSFWorkbook(plugsExcelStream);
        ) {
            for (int i = diaInicio; i <= diaFin; i++) {
                this.actualizarTabla(i, plugsExcelWorkbook);
            }


            try (FileOutputStream plugsOutput = new FileOutputStream(this.RutaPlugsExcel)) {
                plugsExcelWorkbook.write(plugsOutput);
            } catch (FileNotFoundException ex) {
                throw new RuntimeException(ex);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }

            System.out.println("Se cerro el Excel !!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
