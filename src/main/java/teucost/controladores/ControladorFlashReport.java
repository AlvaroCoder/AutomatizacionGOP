package teucost.controladores;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import teucost.modelos.LecturaExcels;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.*;
import java.util.stream.*;

public class ControladorFlashReport {

    private String rutaCarpetaInsumos;
    private String rutaExcelDestino;
    private Boolean extraerDelUltimoArchivo = false;
    private LecturaExcels flashReportExcel;

    // Nombre del mes actual (ej: "MARZO") — se inicializa en el constructor
    private String nombreMesActual;
    // Año actual como String (ej: "2025")
    private String annoActual;

    private final String[] meses = {
            "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
            "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    };

    // Índice de hoja destino
    private final int hojaGMPHBMPH = 0;

    // Columnas donde se pegan los datos (0-based)
    private final int colGMPHSTS = 2;
    private final int colBMPHSTS = 5;
    private final int colGMPHGM  = 7;
    private final int colBMPHGM  = 9;

    // Columnas de mes y año
    private final int colMes  = 0;
    private final int colAnno = 1;

    // Columnas de targets (0-based)
    private final int colTargetAPNGMPHSTS = 3;
    private final int colTargetTPEGMPHSTS = 4;
    private final int colTargetBMPHSTS    = 6;
    private final int colTargetGMPHGM     = 8;
    // NOTA: colTargetBMPHGM comparte índice con colBMPHGM (= 9).
    // Si son columnas distintas, ajusta este valor en tu Excel.
    private final int colTargetBMPHGM     = 10;

    // =========================================================
    //  CONSTRUCTORES
    // =========================================================

    public ControladorFlashReport(String rutaCarpetaInsumos) {
        this.rutaExcelDestino   = "C:\\Users\\alvaro.pupuche\\Desktop\\AUTOMATIZACION KPIS\\RESUMEN EXCELS FLASH REPORTS.xlsx";
        this.rutaCarpetaInsumos = rutaCarpetaInsumos;
        inicializarFechaActual();
    }

    public ControladorFlashReport() {
        this.rutaExcelDestino   = "C:\\Users\\alvaro.pupuche\\Desktop\\AUTOMATIZACION KPIS\\RESUMEN EXCELS FLASH REPORTS.xlsx";
        this.rutaCarpetaInsumos = "C:\\Users\\alvaro.pupuche\\OneDrive - TERMINALES PORTUARIOS EUROANDINOS PAITA S.A\\INSUMOS KPIS\\Flash Report (Lorena)";
        inicializarFechaActual();
    }

    /**
     * Resuelve el mes y año actual en español, en mayúsculas.
     * Ejemplo: nombreMesActual = "MARZO", annoActual = "2025"
     */
    private void inicializarFechaActual() {
        LocalDate hoy = LocalDate.now();
        // Usamos Locale("es") para obtener el nombre en español
        this.nombreMesActual = hoy
                .getMonth()
                .getDisplayName(TextStyle.FULL, new Locale("es"))
                .toUpperCase(new Locale("es"));

        // Normalizar "SETIEMBRE" si el sistema devuelve "SEPTIEMBRE"
        if (this.nombreMesActual.equals("SEPTIEMBRE")) {
            this.nombreMesActual = "SETIEMBRE";
        }

        this.annoActual = String.valueOf(hoy.getYear());
    }

    // =========================================================
    //  PROCESO PRINCIPAL
    // =========================================================

    public void procesarDatosGMPHBMPH() {
        try {
            LocalDate hoy = LocalDate.now();
            this.nombreMesActual = hoy.getMonth()
                    .getDisplayName(TextStyle.FULL, new Locale("es", "PE"))
                    .toUpperCase(Locale.ROOT);
            if (this.nombreMesActual.equals("SEPTIEMBRE")) this.nombreMesActual = "SETIEMBRE";

            // --- 1. Validar carpeta ---
            Path dir = Paths.get(this.rutaCarpetaInsumos).toAbsolutePath().normalize();
            if (!Files.exists(dir) || !Files.isDirectory(dir)) {
                System.out.println("[ERROR] Ruta inválida: " + dir);
                return;
            }

            // --- 2. Listar excels ---
            List<Path> excels = new ArrayList<>();
            if (!extraerDelUltimoArchivo) {
                try (Stream<Path> walk = Files.walk(dir)) {
                    excels = walk
                            .filter(Files::isRegularFile)
                            .filter(this::esExcelCorrecto)
                            .sorted()
                            .collect(Collectors.toList());
                }
                if (excels.isEmpty()) {
                    System.out.println("[INFO] No se encontraron archivos para el mes: " + nombreMesActual);
                    return;
                }
            } else {
                try (Stream<Path> walk = Files.walk(dir)) {
                    excels = walk
                            .filter(Files::isRegularFile)
                            .filter(p -> {
                                String n = p.getFileName().toString().toUpperCase(Locale.ROOT);
                                return n.endsWith(".XLSX") || n.endsWith(".XLS");
                            })
                            .sorted()
                            .collect(Collectors.toList());
                }
                if (excels.isEmpty()) {
                    System.out.println("[INFO] No se encontraron archivos Excel en: " + dir);
                    return;
                }
            }

            // --- 3. Resolver archivo e índice de hoja ---
            Path archivo = extraerDelUltimoArchivo
                    ? excels.get(excels.size() - 1)
                    : excels.get(0);

            System.out.println("archivo = " + archivo.toString());

            int indiceHoja = extraerDelUltimoArchivo
                    ? obtenerIndiceHojaDesdeNombreArchivo(archivo)
                    : hoy.getMonthValue() - 1;

            System.out.println("indiceHoja = " + indiceHoja);

            if (extraerDelUltimoArchivo && indiceHoja == -1) return;

            // --- 4. Leer datos ---
            flashReportExcel = new LecturaExcels(archivo.toString());

            String nombreHoja="January" ;
            switch (indiceHoja){
                case  1:
                    nombreHoja = "February";
                    break;
                case 2:
                    nombreHoja = "March";
                    break;
                case 3:
                    nombreHoja = "April";
                    break;
                case 4:
                    nombreHoja = "May";
                    break;
                default:
                    nombreHoja = "January";
                    break;
            }

            String valorGMPHSTS = flashReportExcel.leerCelda(nombreHoja, 30, 7);
            String valorBMPHSTS = flashReportExcel.leerCelda(nombreHoja ,31, 7);
            String valorGMPHGM  = flashReportExcel.leerCelda(nombreHoja, 42, 7);
            String valorBMPHGM  = flashReportExcel.leerCelda(nombreHoja, 43, 7);
            flashReportExcel.cerrar();

            System.out.println("[INFO] Datos leídos — GMPH STS: " + valorGMPHSTS
                    + " | BMPH STS: " + valorBMPHSTS
                    + " | GMPH GM: "  + valorGMPHGM
                    + " | BMPH GM: "  + valorBMPHGM);

            // --- 5. Pegar en destino ---
            pegarDatosEnDestino(valorGMPHSTS, valorBMPHSTS, valorGMPHGM, valorBMPHGM);

        } catch (Exception err) {
            System.out.println("[ERROR] procesarDatosGMPHBMPH: " + err.getMessage());
            err.printStackTrace();
        }
    }

    /**
     * Obtiene el último Excel de la carpeta (ordenado alfabéticamente),
     * extrae el mes del nombre del archivo y calcula el índice de hoja correspondiente.
     *
     * @return índice de hoja (0=Enero, 1=Febrero, ..., 11=Diciembre), o -1 si hay error
     */
    private int obtenerIndiceHojaDesdeNombreArchivo(Path archivo) {
        String nombreArchivo = archivo.getFileName().toString().toUpperCase(Locale.ROOT);

        String sinExtension = nombreArchivo.contains(".")
                ? nombreArchivo.substring(0, nombreArchivo.lastIndexOf('.'))
                : nombreArchivo;

        String[] partes = sinExtension.split("-");
        if (partes.length < 1) {
            System.out.println("[ERROR] Nombre de archivo no reconocido: " + nombreArchivo);
            return -1;
        }

        String mesDelArchivo = partes[0].trim();

        Map<String, Integer> mapaIndices = new LinkedHashMap<>();
        mapaIndices.put("ENERO",      0);
        mapaIndices.put("FEBRERO",    1);
        mapaIndices.put("MARZO",      2);
        mapaIndices.put("ABRIL",      3);
        mapaIndices.put("MAYO",       4);
        mapaIndices.put("JUNIO",      5);
        mapaIndices.put("JULIO",      6);
        mapaIndices.put("AGOSTO",     7);
        mapaIndices.put("SETIEMBRE",  8);
        mapaIndices.put("SEPTIEMBRE", 8);
        mapaIndices.put("OCTUBRE",    9);
        mapaIndices.put("NOVIEMBRE",  10);
        mapaIndices.put("DICIEMBRE",  11);

        Integer indice = mapaIndices.get(mesDelArchivo);
        if (indice == null) {
            System.out.println("[ERROR] Mes no reconocido: " + mesDelArchivo);
            return -1;
        }

        this.nombreMesActual = mesDelArchivo;
        System.out.println("[INFO] Archivo: " + nombreArchivo
                + " → Mes: " + mesDelArchivo + " → Hoja: " + indice);
        return indice;
    }

    // =========================================================
    //  ESCRITURA EN DESTINO — PRESERVANDO FORMATO DE TABLA
    // =========================================================

    /**
     * Abre el Excel destino, localiza la última fila con datos,
     * copia el estilo de la fila anterior (mes anterior) y escribe
     * la nueva fila con mes, año, valores y targets.
     */
    private void pegarDatosEnDestino(String gmphSTS, String bmphSTS,
                                     String gmphGM,  String bmphGM) throws IOException {

        File archivoDestino = new File(rutaExcelDestino);
        if (!archivoDestino.exists()) {
            System.out.println("[ERROR] No existe el archivo destino: " + rutaExcelDestino);
            return;
        }

        // Abrir el workbook destino
        LecturaExcels excelDestino = new LecturaExcels(rutaExcelDestino);

        // --- Obtener la fila anterior (último mes) para copiar targets ---
        // colMes = 0, sirve como columna de referencia para detectar datos
        Row filaAnterior = excelDestino.obtenerUltimaFilaConDatosEnTabla(
                hojaGMPHBMPH, 0, colMes);

        // --- Insertar nueva fila al final de la tabla (expande + copia estilos) ---
        Row nuevaFila = excelDestino.insertarFilaEnTabla(hojaGMPHBMPH, 0);

        // --- Escribir MES y AÑO ---
        setCeldaTexto(nuevaFila, colMes,  nombreMesActual);
        setCeldaTexto(nuevaFila, colAnno, annoActual);

        // --- Escribir valores KPI ---
        setCeldaNumero(nuevaFila, colGMPHSTS, parsearDouble(gmphSTS));
        setCeldaNumero(nuevaFila, colBMPHSTS, parsearDouble(bmphSTS));
        setCeldaNumero(nuevaFila, colGMPHGM,  parsearDouble(gmphGM));
        setCeldaNumero(nuevaFila, colBMPHGM,  parsearDouble(bmphGM));

        // --- Targets: valor por defecto = valor del mes anterior ---
        copiarTargetDesdeFila(filaAnterior, nuevaFila, colTargetAPNGMPHSTS);
        copiarTargetDesdeFila(filaAnterior, nuevaFila, colTargetTPEGMPHSTS);
        copiarTargetDesdeFila(filaAnterior, nuevaFila, colTargetBMPHSTS);
        copiarTargetDesdeFila(filaAnterior, nuevaFila, colTargetGMPHGM);
        copiarTargetDesdeFila(filaAnterior, nuevaFila, colTargetBMPHGM);

        // --- Guardar ---
        excelDestino.guardar();
        excelDestino.cerrar();

        System.out.println("[OK] Nueva fila insertada en tabla → " + rutaExcelDestino);
    }

    // =========================================================
    //  UTILIDADES INTERNAS
    // =========================================================

    /**
     * Recorre la columna de MES para encontrar la primera fila sin datos.
     * Parte desde la fila 1 (salta encabezado en fila 0).
     */
    private int encontrarPrimeraFilaVacia(Sheet hoja) {
        int ultima = hoja.getLastRowNum();
        for (int f = 1; f <= ultima; f++) {
            Row row = hoja.getRow(f);
            if (row == null) return f;
            Cell celda = row.getCell(colMes);
            if (celda == null || celda.getCellType() == CellType.BLANK
                    || celda.toString().trim().isEmpty()) {
                return f;
            }
        }
        return ultima + 1; // Agregar al final si todas tienen datos
    }

    /**
     * Copia el valor numérico de un target desde la fila anterior.
     * Si no existe o no es número, deja la celda vacía (estilo ya fue copiado).
     */
    private void copiarTargetDesdeFila(Row filaAnterior, Row nuevaFila, int columna) {
        if (filaAnterior == null) return;
        Cell origen  = filaAnterior.getCell(columna);
        Cell destino = nuevaFila.getCell(columna);
        if (destino == null) destino = nuevaFila.createCell(columna);

        if (origen != null && origen.getCellType() == CellType.NUMERIC) {
            destino.setCellValue(origen.getNumericCellValue());
        } else if (origen != null && origen.getCellType() == CellType.STRING) {
            String val = origen.getStringCellValue().trim();
            if (!val.isEmpty()) {
                try {
                    destino.setCellValue(Double.parseDouble(val.replace(",", ".")));
                } catch (NumberFormatException e) {
                    destino.setCellValue(val);
                }
            }
        }
    }

    /** Escribe texto en una celda, sin perder el estilo previamente asignado. */
    private void setCeldaTexto(Row fila, int columna, String valor) {
        Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor);
    }

    /** Escribe número en una celda, sin perder el estilo previamente asignado. */
    private void setCeldaNumero(Row fila, int columna, double valor) {
        Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor);
    }

    /** Convierte String a double de forma segura; retorna 0.0 si falla. */
    private double parsearDouble(String valor) {
        if (valor == null || valor.trim().isEmpty()) return 0.0;
        try {
            return Double.parseDouble(valor.trim().replace(",", "."));
        } catch (NumberFormatException e) {
            System.out.println("[WARN] No se pudo parsear como número: '" + valor + "'");
            return 0.0;
        }
    }

    // =========================================================
    //  VALIDACIÓN DE ARCHIVOS
    // =========================================================

    /**
     * Verifica que el archivo sea un Excel cuyo nombre empiece con
     * el mes actual en español. Formato esperado: "MARZO - ..." o "MARZO.xlsx"
     *
     * CORRECCIÓN: Se cambió allMatch → anyMatch para que funcione correctamente.
     */
    public boolean esExcelCorrecto(Path p) {
        String f = p.getFileName().toString().toUpperCase(Locale.ROOT);
        // Quitar extensión y separar por "-"
        String sinExtension = f.contains(".") ? f.substring(0, f.lastIndexOf('.')) : f;
        String[] partes = sinExtension.split("-");

        if (partes.length < 2) return false;

        String mesCandidato  = partes[0].trim();
        String annoCandidato = partes[1].trim();

        String mesActual  = LocalDate.now()
                .getMonth()
                .getDisplayName(TextStyle.FULL, new Locale("es", "PE"))
                .toUpperCase(Locale.ROOT);

        // Caso especial: Java devuelve "SEPTIEMBRE", tú usas "SETIEMBRE"
        if (mesActual.equals("SEPTIEMBRE")) mesActual = "SETIEMBRE";

        String annoActual = String.valueOf(LocalDate.now().getYear());

        return mesCandidato.equals(mesActual) && annoCandidato.equals(annoActual);
    }

    // =========================================================
    //  GETTERS / SETTERS (para tests o modificación de targets)
    // =========================================================

    public String getNombreMesActual() { return nombreMesActual; }
    public String getAnnoActual()      { return annoActual; }
    public String getRutaExcelDestino(){ return rutaExcelDestino; }

    public void setRutaExcelDestino(String ruta) { this.rutaExcelDestino = ruta; }
    public void setRutaCarpetaInsumos(String ruta){ this.rutaCarpetaInsumos = ruta; }
    public void setExtraerDelUltimoArchivo(Boolean extraer) {this.extraerDelUltimoArchivo = extraer; }
}
