package teucost.controladores;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import teucost.modelos.LecturaExcels;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ControladorTeus {
    // Al inicio de la clase, junto a los otros atributos:
    private final String[] meses = {
            "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
            "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    };

    private String rutaCarpetaOM = "C:\\Users\\alvaro.pupuche\\OneDrive - TERMINALES PORTUARIOS EUROANDINOS PAITA S.A\\INSUMOS KPIS\\OM (Finanzas)";

    private String nombreMesActual;
    // Año actual como String (ej: "2025")
    private String annoActual;

    private String rutaExcelThroughput;
    private String rutaExcelDestino;
    private int nroMesSeleccionado = 1;

    private final int colAnno = 0;
    private final int colMes = 1;

    private final int colTeus = 2;
    private final int colBudget = 3;

    // Hoja del Throughput donde extraigo la información
    private final String nombreHojaTrhoughput = "Flash Report";
    private final int filaTotalMes = 6;
    private final int filaTotalResultado = 39;

    private final int hojaDestino = 0;

    private final int filaNombreMes = 4;
    private final int filaBudget = 6;

    public ControladorTeus(){
        inicializarFechaActual();
    }

    public ControladorTeus(String rutaExcelThroughput, String rutaExcelDestino){
        this.rutaExcelDestino = rutaExcelDestino;
        this.rutaExcelThroughput = rutaExcelThroughput;
        inicializarFechaActual();
    }

    private void inicializarFechaActual() {
        // Solo inicializa el año — no depende de nroMesSeleccionado
        this.annoActual = String.valueOf(LocalDate.now().getYear());
    }

    private String obtenerNombreMes(int nroMes) {
        String nombre = LocalDate.of(LocalDate.now().getYear(), nroMes, 1)
                .getMonth()
                .getDisplayName(TextStyle.FULL, new Locale("es"))
                .toUpperCase(new Locale("es"));

        return nombre.equals("SEPTIEMBRE") ? "SETIEMBRE" : nombre;
    }

    public List<String> extraerListaBudgets() {
        try {
            // --- 1. Validar carpeta raíz OM ---
            Path carpetaOM = Paths.get(this.rutaCarpetaOM).toAbsolutePath().normalize();
            if (!Files.exists(carpetaOM) || !Files.isDirectory(carpetaOM)) {
                System.out.println("[ERROR] Ruta inválida: " + carpetaOM);
                return Collections.emptyList(); // ✅ corregido
            }

            // --- 2. Entrar a la carpeta del año actual (ej: "2026") ---
            Path carpetaAnno = carpetaOM.resolve(this.annoActual);
            if (!Files.exists(carpetaAnno) || !Files.isDirectory(carpetaAnno)) {
                System.out.println("[ERROR] No existe la carpeta del año: " + carpetaAnno);
                return Collections.emptyList(); // ✅ corregido
            }

            System.out.println("[INFO] Buscando en carpeta: " + carpetaAnno);

            String patronEsperado = "OM-ENERO_" + this.annoActual;

            List<Path> archivosOM;
            try (Stream<Path> walk = Files.walk(carpetaAnno, 1)) {
                archivosOM = walk
                        .filter(Files::isRegularFile)
                        .filter(p -> {
                            String nombre = p.getFileName().toString().toUpperCase(Locale.ROOT);
                            String sinExtension = nombre.contains(".")
                                    ? nombre.substring(0, nombre.lastIndexOf('.'))
                                    : nombre;
                            return sinExtension.equals(patronEsperado)
                                    && (nombre.endsWith(".XLSX") || nombre.endsWith(".XLS") || nombre.endsWith(".XLSM"));
                        })
                        .collect(Collectors.toList());
            }

            if (archivosOM.isEmpty()) {
                System.out.println("[ERROR] No se encontró el archivo '" + patronEsperado
                        + "' en: " + carpetaAnno);
                return Collections.emptyList(); // ✅ corregido
            }

            Path archivoEnero = archivosOM.get(0);
            System.out.println("[INFO] Archivo encontrado: " + archivoEnero.getFileName());

            LecturaExcels omExcel = new LecturaExcels(archivoEnero.toString());
            int colInicio = 17; // Columna R
            int colFin    = 28; // Columna AC

            List<String> valoresFila = omExcel.leerFilaEnRango(
                    "Monthly Traffic", filaBudget, colInicio, colFin);
            omExcel.cerrar();

            System.out.println("[INFO] Budgets extraídos (" + valoresFila.size() + " valores): " + valoresFila);

            return valoresFila;

        } catch (Exception err) {
            System.out.println("[ERROR] extraerListaBudgets: " + err.getMessage());
            err.printStackTrace();
            return Collections.emptyList(); // ✅ corregido
        }
    }

    private String obtenerLetraColumna(int indice) {
        StringBuilder letra = new StringBuilder();
        indice++; // Pasar a 1-based
        while (indice > 0) {
            indice--;
            letra.insert(0, (char) ('A' + indice % 26));
            indice /= 26;
        }
        return letra.toString();
    }

    public void procesarDatosTeus() {
        try {
            this.nombreMesActual = obtenerNombreMes(nroMesSeleccionado);

            // --- 0. Extraer lista de budgets y seleccionar el del mes actual ---
            // nroMesSeleccionado=1 (Enero) → índice 0, =2 (Febrero) → índice 1, etc.
            List<String> listaBudgets = extraerListaBudgets();
            int indiceMes = nroMesSeleccionado - 1;

            String valorBudgetMes = "0";
            if (!listaBudgets.isEmpty() && indiceMes < listaBudgets.size()) {
                valorBudgetMes = listaBudgets.get(indiceMes);
                System.out.println("[INFO] Budget para " + nombreMesActual
                        + " (índice " + indiceMes + "): " + valorBudgetMes);
            } else {
                System.out.println("[WARN] No se encontró budget para el índice " + indiceMes
                        + " — se usará 0 por defecto");
            }

            LecturaExcels flashReportThroughput = new LecturaExcels(this.rutaExcelThroughput);

            String textoTotalBuscado = "TOTAL " + meses[nroMesSeleccionado - 1];
            System.out.println("[INFO] Buscando columna con texto: '" + textoTotalBuscado + "'");

            List<String> filaTotalesMes = flashReportThroughput.leerFila(
                    this.nombreHojaTrhoughput, this.filaTotalMes);

            int columnaEncontrada = -1;
            for (int i = 0; i < filaTotalesMes.size(); i++) {
                if (filaTotalesMes.get(i).trim().toUpperCase(Locale.ROOT).equals(textoTotalBuscado)) {
                    columnaEncontrada = i;
                    break;
                }
            }

            if (columnaEncontrada == -1) {
                System.out.println("[ERROR] No se encontró columna '" + textoTotalBuscado + "'");
                System.out.println("[DEBUG] Fila " + filaTotalMes + ": " + filaTotalesMes);
                flashReportThroughput.cerrar();
                return;
            }

            String valorTeus = flashReportThroughput.leerCelda(
                    this.nombreHojaTrhoughput, this.filaTotalResultado, columnaEncontrada);

            flashReportThroughput.cerrar();

            // ✅ Pasar el budget extraído al método de pegado
            pegarDatosEnDestino(valorTeus, valorBudgetMes);

            System.out.println("==============================================");
            System.out.println(" Mes       : " + nombreMesActual);
            System.out.println(" TEUs      : " + valorTeus);
            System.out.println(" Budget    : " + valorBudgetMes);
            System.out.println("==============================================");

        } catch (Exception err) {
            System.out.println("[ERROR] procesarDatosTeus: " + err.getMessage());
            err.printStackTrace();
        }
    }

    private void pegarDatosEnDestino(String valorTeus, String valorBudget) throws IOException {
        File archivoDestino = new File(this.rutaExcelDestino);
        if (!archivoDestino.exists()) {
            System.out.println("[ERROR] No existe el archivo destino: " + rutaExcelDestino);
            return;
        }

        LecturaExcels excelDestino = new LecturaExcels(rutaExcelDestino);

        Row nuevaFila = excelDestino.insertarFilaEnTabla(hojaDestino, 0);

        setCeldaTexto(nuevaFila,   colMes,    nombreMesActual);
        setCeldaTexto(nuevaFila,   colAnno,   annoActual);
        setCeldaNumero(nuevaFila,  colTeus,   parsearDouble(valorTeus));
        setCeldaNumero(nuevaFila,  colBudget, parsearDouble(valorBudget)); // ✅ budget real

        excelDestino.guardar();
        excelDestino.cerrar();

        System.out.println("[OK] Datos pegados correctamente en: " + rutaExcelDestino);
    }

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
    private double parsearDouble(String valor) {
        if (valor == null || valor.trim().isEmpty()) return 0.0;
        try {
            return Double.parseDouble(valor.trim().replace(",", "."));
        } catch (NumberFormatException e) {
            System.out.println("[WARN] No se pudo parsear como número: '" + valor + "'");
            return 0.0;
        }
    }


    private void setCeldaTexto(Row fila, int columna, String valor) {
        Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor);
    }

    private void setCeldaNumero(Row fila, int columna, double valor) {
        Cell celda = fila.getCell(columna);
        if (celda == null) celda = fila.createCell(columna);
        celda.setCellValue(valor);
    }


    public void setNroMesSeleccionado(int nroMesSeleccionado){
        this.nroMesSeleccionado = nroMesSeleccionado;
    }

    public void setRutaExcelThroughput(String rutaExcelThroughput) {
        this.rutaExcelThroughput = rutaExcelThroughput;
    }

    public void setRutaExcelDestino(String rutaExcelDestino) {
        this.rutaExcelDestino = rutaExcelDestino;
    }
}

