package cuadrillas;

import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExtraccionCuadrillas {
    private static final String TESSDATA_PATH = "C:\\Users\\alvaro.pupuche\\AppData\\Local\\Programs\\Tesseract-OCR\\tessdata";
    private static final String IDIOMA = "spa+eng";
    private static final float DPI = 300f;

    private String rutaCarpeta;
    public String rutaDocumentoEntrada;
    public String rutaCarpetaSalida;

    public ExtraccionCuadrillas() {}

    public ExtraccionCuadrillas(String rutaCarpeta) {
        this.rutaCarpeta = rutaCarpeta;
    }

    public ExtraccionCuadrillas(String rutaDocumentoEntrada, String rutaCarpetaSalida) {
        this.rutaDocumentoEntrada = rutaDocumentoEntrada;
        this.rutaCarpetaSalida = rutaCarpetaSalida;
    }

    // ─────────────────────────────────────────────
    // EXTRAER TEXTO CON OCR (ya existente)
    // ─────────────────────────────────────────────

    public static Map<Integer, String> extraerTextoPdf(String rutaPdf) throws IOException, TesseractException {
        System.out.println("\n📄 Procesando: " + rutaPdf);
        System.out.println("   Idioma: " + IDIOMA + " | DPI: " + DPI);

        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath(TESSDATA_PATH);
        tesseract.setLanguage(IDIOMA);
        tesseract.setPageSegMode(6);
        tesseract.setOcrEngineMode(1);

        Map<Integer, String> resultados = new LinkedHashMap<>();

        try (PDDocument documento = PDDocument.load(new File(rutaPdf))) {
            PDFRenderer renderer = new PDFRenderer(documento);
            int totalPaginas = documento.getNumberOfPages();
            System.out.println("   Total de páginas: " + totalPaginas);

            for (int i = 0; i < totalPaginas; i++) {
                System.out.println("\n🔍 Analizando página " + (i + 1) + "/" + totalPaginas + "...");
                BufferedImage imagen = renderer.renderImageWithDPI(i, DPI, ImageType.RGB);
                String texto = tesseract.doOCR(imagen);
                resultados.put(i + 1, texto);
                System.out.println("   ✅ Página " + (i + 1) + " completada (" + texto.length() + " caracteres)");
            }
        }

        System.out.println("\n✅ OCR completado. " + resultados.size() + " página(s) procesada(s).");
        return resultados;
    }

    // ─────────────────────────────────────────────
    // GUARDAR COMO TXT (ya existente)
    // ─────────────────────────────────────────────

    public static void guardarComoTxt(String rutaPdf, Map<Integer, String> resultado) throws IOException {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String nombreBase = Paths.get(rutaPdf).getFileName().toString().replace(".pdf", "");
        String rutaSalida = nombreBase + "_" + timestamp + ".txt";

        StringBuilder contenido = new StringBuilder();
        contenido.append("Archivo: ").append(rutaPdf).append("\n");
        contenido.append("Páginas: ").append(resultado.size()).append("\n\n");

        String separador = new String(new char[50]).replace("\0", "=");
        resultado.forEach((pagina, texto) -> {
            contenido.append(separador).append("\n");
            contenido.append("PÁGINA ").append(pagina).append("\n");
            contenido.append(separador).append("\n");
            contenido.append(texto).append("\n\n");
        });

        try (BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(new FileOutputStream(rutaSalida), "UTF-8"))) {
            writer.write(contenido.toString());
        }
        System.out.println("💾 TXT guardado en: " + rutaSalida);
    }

    // ─────────────────────────────────────────────
    // GUARDAR COMO EXCEL CRUDO (ya existente)
    // ─────────────────────────────────────────────

    public static void guardarComoExcel(String rutaPdf, Map<Integer, String> resultado) throws IOException {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String nombreBase = Paths.get(rutaPdf).getFileName().toString().replace(".pdf", "");
        String rutaSalida = nombreBase + "_" + timestamp + ".xlsx";

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet hoja = workbook.createSheet("Texto extraído");

            Row encabezado = hoja.createRow(0);
            CellStyle estiloEncabezado = workbook.createCellStyle();
            Font fuenteNegrita = workbook.createFont();
            fuenteNegrita.setBold(true);
            estiloEncabezado.setFont(fuenteNegrita);

            encabezado.createCell(0).setCellValue("Página");
            encabezado.createCell(1).setCellValue("Texto extraído");
            encabezado.getCell(0).setCellStyle(estiloEncabezado);
            encabezado.getCell(1).setCellStyle(estiloEncabezado);

            int fila = 1;
            for (Map.Entry<Integer, String> entry : resultado.entrySet()) {
                Row row = hoja.createRow(fila++);
                row.createCell(0).setCellValue(entry.getKey());
                Cell celdaTexto = row.createCell(1);
                celdaTexto.setCellValue(entry.getValue());
                row.setHeightInPoints(60);
            }

            hoja.setColumnWidth(0, 2500);
            hoja.setColumnWidth(1, 20000);

            try (FileOutputStream fos = new FileOutputStream(rutaSalida)) {
                workbook.write(fos);
            }
        }
        System.out.println("📊 Excel guardado en: " + rutaSalida);
    }

    // ─────────────────────────────────────────────
    // PROCESAR CARPETA COMPLETA (ya existente)
    // ─────────────────────────────────────────────

    public static void procesarCarpeta(String carpetaEntrada, String carpetaSalida) throws IOException {
        Files.createDirectories(Paths.get(carpetaSalida));

        List<Path> pdfs = Files.list(Paths.get(carpetaEntrada))
                .filter(p -> p.toString().toLowerCase().endsWith(".pdf"))
                .sorted()
                .collect(Collectors.toList());

        if (pdfs.isEmpty()) {
            System.out.println("⚠️  No se encontraron PDFs en: " + carpetaEntrada);
            return;
        }

        System.out.println("\n📂 Se encontraron " + pdfs.size() + " PDFs en '" + carpetaEntrada + "'");

        for (Path pdf : pdfs) {
            try {
                Map<Integer, String> resultado = extraerTextoPdf(pdf.toString());
                guardarComoTxt(carpetaSalida + "/" + pdf.getFileName(), resultado);
                guardarComoExcel(carpetaSalida + "/" + pdf.getFileName(), resultado);
            } catch (Exception e) {
                System.err.println("❌ Error procesando " + pdf.getFileName() + ": " + e.getMessage());
            }
        }

        System.out.println("\n✅ Todos los PDFs procesados. Resultados en: '" + carpetaSalida + "'");
    }

    // ─────────────────────────────────────────────
    // NUEVO: EXTRAER DATOS SAP Y PEGAR EN EXCEL
    // ─────────────────────────────────────────────

    /**
     * Lee todos los PDFs de una carpeta, extrae la fila de datos SAP
     * y la escribe como fila en el Excel destino indicado.
     *
     * Uso:
     *   ExtraccionCuadrillas ec = new ExtraccionCuadrillas();
     *   ec.extraerDatosAExcel(
     *       "C:/mis_pdfs",
     *       "C:/resultado.xlsx"
     *   );
     *
     * @param rutaCarpetaPdfs  Carpeta con los PDFs escaneados
     * @param rutaExcelDestino Ruta del Excel donde se pegarán los datos (.xlsx)
     */
    public void extraerDatosAExcel(String rutaCarpetaPdfs, String rutaExcelDestino) {

        Path carpetaPdfs  = Paths.get(rutaCarpetaPdfs).toAbsolutePath().normalize();
        Path archivoExcel = Paths.get(rutaExcelDestino).toAbsolutePath().normalize();

        // Validar carpeta origen
        if (!Files.exists(carpetaPdfs) || !Files.isDirectory(carpetaPdfs)) {
            System.out.println("[ERROR] Carpeta de PDFs inválida: " + carpetaPdfs);
            return;
        }

        // Buscar todos los PDFs en la carpeta
        List<Path> pdfs;
        try (Stream<Path> walk = Files.walk(carpetaPdfs)) {
            pdfs = walk
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString()
                            .toUpperCase(Locale.ROOT).endsWith(".PDF"))
                    .sorted()
                    .collect(Collectors.toList());
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo leer la carpeta: " + e.getMessage());
            return;
        }

        if (pdfs.isEmpty()) {
            System.out.println("[INFO] No se encontraron PDFs en: " + carpetaPdfs);
            return;
        }

        System.out.println("\n════════ EXTRACCIÓN A EXCEL ════════");
        System.out.println("PDFs encontrados : " + pdfs.size());
        System.out.println("Excel destino    : " + archivoExcel);

        // Cargar Excel existente o crear uno nuevo
        Workbook workbook;
        Sheet hoja;

        if (Files.exists(archivoExcel)) {
            try (FileInputStream fis = new FileInputStream(archivoExcel.toFile())) {
                workbook = new XSSFWorkbook(fis);
                hoja = workbook.getSheetAt(0);
                System.out.println("[INFO] Excel existente cargado. Filas actuales: " + hoja.getLastRowNum());
            } catch (Exception e) {
                System.out.println("[ERROR] No se pudo abrir el Excel: " + e.getMessage());
                return;
            }
        } else {
            workbook = new XSSFWorkbook();
            hoja = workbook.createSheet("Datos SAP");
            crearEncabezadosExcel(hoja, workbook);
            System.out.println("[INFO] Excel nuevo creado con encabezados.");
        }

        // Configurar Tesseract
        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath(TESSDATA_PATH);
        tesseract.setLanguage(IDIOMA);
        tesseract.setPageSegMode(6);
        tesseract.setOcrEngineMode(1);

        int filaActual = hoja.getLastRowNum() + 1;
        int procesados = 0;
        int errores    = 0;

        // Procesar cada PDF
        for (Path pdf : pdfs) {
            System.out.println("\n[PROCESANDO] " + pdf.getFileName());
            try {
                // Paso 1: OCR → texto completo
                String textoCompleto = extraerTextoConOCR(pdf, tesseract);

                if (textoCompleto == null || textoCompleto.trim().isEmpty()) {
                    System.out.println("[SKIP] Sin texto reconocible: " + pdf.getFileName());
                    errores++;
                    continue;
                }

                // Paso 2: Parsear fila de la tabla SAP
                List<String> datosFila = parsearFilaSAP(textoCompleto, pdf.getFileName().toString());

                if (datosFila == null) {
                    System.out.println("[SKIP] No se encontró tabla SAP en: " + pdf.getFileName());
                    errores++;
                    continue;
                }

                // Paso 3: Escribir fila en el Excel
                Row fila = hoja.createRow(filaActual++);
                for (int col = 0; col < datosFila.size(); col++) {
                    fila.createCell(col).setCellValue(datosFila.get(col));
                }

                System.out.println("[OK] Fila agregada: " + pdf.getFileName());
                procesados++;

            } catch (Exception e) {
                System.out.println("[ERROR] " + pdf.getFileName() + " → " + e.getMessage());
                errores++;
            }
        }

        // Guardar el Excel
        try (FileOutputStream fos = new FileOutputStream(archivoExcel.toFile())) {
            workbook.write(fos);
            workbook.close();
            System.out.println("\n════════ RESUMEN ════════");
            System.out.println("✅ Filas agregadas : " + procesados);
            System.out.println("❌ Errores         : " + errores);
            System.out.println("💾 Excel guardado  : " + archivoExcel);
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo guardar el Excel: " + e.getMessage());
        }
    }

    // ─────────────────────────────────────────────
    // NUEVO: EXTRAER TEXTO DE UN PDF CON OCR
    // ─────────────────────────────────────────────

    private String extraerTextoConOCR(Path rutaPdf, Tesseract tesseract) throws Exception {
        StringBuilder sb = new StringBuilder();
        try (PDDocument doc = PDDocument.load(rutaPdf.toFile())) {
            PDFRenderer renderer = new PDFRenderer(doc);
            for (int i = 0; i < doc.getNumberOfPages(); i++) {
                BufferedImage img = renderer.renderImageWithDPI(i, DPI, ImageType.RGB);
                sb.append(tesseract.doOCR(img));
            }
        }
        return sb.toString();
    }

    // ─────────────────────────────────────────────
    // NUEVO: PARSEAR FILA DE DATOS SAP DEL TEXTO OCR
    // ─────────────────────────────────────────────

    /**
     * Extrae los campos de la tabla SAP del texto OCR.
     * Busca el bloque entre "Referencia" (fin del encabezado)
     * y "Solicitado" (inicio de firmas).
     *
     * Columnas en orden:
     * Archivo PDF | # | Descripcion | Codigo | Almacen | Medida |
     * Precio Unitario | Cantidad | Precio Total | Proveedor | Cuenta |
     * Proyecto | Centro Costo 1 | Centro Costo 2 | Centro Costo 3 |
     * Centro Costo 4 | Centro Costo 5 | Referencia
     */
    private List<String> parsearFilaSAP(String texto, String nombreArchivo) {

        // Normalizar saltos de línea
        String textoNorm = texto.replace("\r\n", "\n").replace("\r", "\n");

        // Localizar el bloque de datos entre encabezado y sección de firmas
        String marcadorInicio = "Referencia";  // última columna del encabezado
        String marcadorFin    = "Solicitado";  // inicio del pie del formulario

        int idxInicio = textoNorm.indexOf(marcadorInicio);
        int idxFin    = textoNorm.indexOf(marcadorFin);

        if (idxInicio == -1 || idxFin == -1 || idxInicio >= idxFin) {
            System.out.println("   [WARN] Marcadores no encontrados en: " + nombreArchivo);
            return null;
        }

        // Texto entre "Referencia\n" y "Solicitado"
        String bloqueDatos = textoNorm
                .substring(idxInicio + marcadorInicio.length(), idxFin)
                .trim();

        if (bloqueDatos.isEmpty()) {
            System.out.println("   [WARN] Bloque de datos vacío en: " + nombreArchivo);
            return null;
        }

        // Unir todo el bloque en una sola línea normalizada
        String lineaCompleta = bloqueDatos
                .replaceAll("[\r\n]+", " ")  // saltos de línea → espacio
                .replaceAll("\\s+", " ")      // múltiples espacios → uno solo
                .trim();

        // ── Estrategia de extracción por campos conocidos ──────────────────
        //
        // El texto OCR siempre tiene este patrón (con posibles variaciones):
        //   1  DESCRIPCION  CODIGO  ALMACEN  MEDIDA  PrecioUnit  Cantidad  PrecioTotal
        //   PROVEEDOR  CUENTA  PROYECTO  CC1  CC2  CC3  CC4  CC5  REFERENCIA
        //
        // Los 3 valores numéricos consecutivos son: PrecioUnitario, Cantidad, PrecioTotal
        // Los identificamos con regex: tres grupos "número decimal" seguidos
        // Ejemplo: "0.00 2.00 0.00" o "0,00 2,00 0,00"
        // ───────────────────────────────────────────────────────────────────

        // Regex: busca exactamente 3 decimales consecutivos separados por espacio
        java.util.regex.Pattern patNumeros = java.util.regex.Pattern.compile(
                "(\\d+[.,]\\d+)\\s+(\\d+[.,]\\d+)\\s+(\\d+[.,]\\d+)"
        );
        java.util.regex.Matcher matcherNum = patNumeros.matcher(lineaCompleta);

        String precioUnitario = "-";
        String cantidad       = "-";
        String precioTotal    = "-";
        int    posInicioNum   = -1;
        int    posFinNum      = -1;

        if (matcherNum.find()) {
            precioUnitario = matcherNum.group(1);
            cantidad       = matcherNum.group(2);  // ← CANTIDAD siempre es el 2do número
            precioTotal    = matcherNum.group(3);
            posInicioNum   = matcherNum.start();
            posFinNum      = matcherNum.end();
            System.out.println("   [CANTIDAD detectada] " + cantidad
                    + " | PrecioUnit: " + precioUnitario
                    + " | PrecioTotal: " + precioTotal);
        } else {
            System.out.println("   [WARN] No se encontraron valores numéricos en: " + nombreArchivo);
        }

        // ── Texto ANTES de los números: contiene #, Descripción, Código, Almacén, Medida
        String textoAntes = posInicioNum > 0
                ? lineaCompleta.substring(0, posInicioNum).trim()
                : lineaCompleta;

        // ── Texto DESPUÉS de los números: Proveedor, Cuenta, Proyecto, CCs, Referencia
        String textoDespues = posFinNum > 0 && posFinNum < lineaCompleta.length()
                ? lineaCompleta.substring(posFinNum).trim()
                : "";

        // Tokenizar la parte anterior por espacios múltiples o simple espacio
        String[] tokensAntes   = textoAntes.isEmpty()   ? new String[0] : textoAntes.split("\\s{2,}");
        String[] tokensDespues = textoDespues.isEmpty() ? new String[0] : textoDespues.split("\\s+|-\\s+");

        // Si tokensAntes tiene muy pocos elementos, intentar con espacio simple
        if (tokensAntes.length < 3) {
            tokensAntes = textoAntes.split("\\s+");
        }

        // Extraer campos de la parte anterior
        // Orden esperado: # | Descripcion (puede ser multi-palabra) | Codigo | Almacen | Medida
        String numero      = tokensAntes.length > 0 ? tokensAntes[0] : "-";
        String codigo      = "-";
        String almacen     = "-";
        String medida      = "-";
        String descripcion = "-";

        // El código es numérico largo (ej: 713200001) — buscarlo por patrón
        java.util.regex.Pattern patCodigo = java.util.regex.Pattern.compile("\\b(\\d{6,12})\\b");
        java.util.regex.Matcher matcherCod = patCodigo.matcher(textoAntes);
        if (matcherCod.find()) {
            codigo = matcherCod.group(1);
            // Descripción = texto antes del código
            descripcion = textoAntes.substring(0, matcherCod.start()).replaceAll("^\\d+\\s*", "").trim();
            // Almacén = texto después del código (lo que queda antes de los números)
            String restoAntes = textoAntes.substring(matcherCod.end()).trim();
            if (!restoAntes.isEmpty()) {
                almacen = restoAntes;
            }
        } else {
            // Sin código numérico claro: tomar tokens directamente
            if (tokensAntes.length > 1) descripcion = tokensAntes[1];
            if (tokensAntes.length > 2) codigo      = tokensAntes[2];
            if (tokensAntes.length > 3) almacen     = tokensAntes[3];
            if (tokensAntes.length > 4) medida       = tokensAntes[4];
        }

        // Extraer campos de la parte posterior
        // Orden esperado: Proveedor Cuenta Proyecto CC1 CC2 CC3 CC4 CC5 Referencia
        // Los guiones "-" son valores vacíos en el formulario SAP
        List<String> camposDespues = new ArrayList<>();
        if (!textoDespues.isEmpty()) {
            for (String t : textoDespues.split("\\s+")) {
                String tk = t.trim().replaceAll("[,.]$", "");
                if (!tk.isEmpty()) camposDespues.add(tk);
            }
        }

        // Rellenar con "-" hasta tener 9 campos (Proveedor→Referencia)
        while (camposDespues.size() < 9) camposDespues.add("-");

        // Separar "TOTAL 0.00" del final si el OCR lo pegó al bloque
        // (ya fue filtrado arriba pero por seguridad)
        List<String> fila = new ArrayList<>();
        fila.add(nombreArchivo);   // col 0: trazabilidad
        fila.add(numero);          // col 1: #
        fila.add(descripcion);     // col 2: Descripcion
        fila.add(codigo);          // col 3: Codigo
        fila.add(almacen);         // col 4: Almacen
        fila.add(medida);          // col 5: Medida
        fila.add(precioUnitario);  // col 6: Precio Unitario
        fila.add(cantidad);        // col 7: Cantidad  ← campo clave
        fila.add(precioTotal);     // col 8: Precio Total
        // cols 9-17: Proveedor, Cuenta, Proyecto, CC1, CC2, CC3, CC4, CC5, Referencia
        for (int i = 0; i < 9; i++) {
            fila.add(camposDespues.get(i));
        }

        return fila;
    }

    // ─────────────────────────────────────────────
    // NUEVO: EXTRAER TEXTO DE PDFs Y GUARDAR COMO TXT
    // ─────────────────────────────────────────────

    /**
     * Lee todos los PDFs de una carpeta, aplica OCR y guarda
     * el texto extraído como archivos .txt en la carpeta destino.
     * Un .txt por cada PDF, con el mismo nombre base.
     *
     * Uso:
     *   ExtraccionCuadrillas ec = new ExtraccionCuadrillas();
     *   ec.extraerPdfsATxt(
     *       "C:/mis_pdfs",
     *       "C:/mis_textos"
     *   );
     *
     * @param rutaCarpetaPdfs  Carpeta con los PDFs escaneados
     * @param rutaCarpetaTxt   Carpeta destino donde se guardarán los .txt
     */
    public void extraerPdfsATxt(String rutaCarpetaPdfs, String rutaCarpetaTxt) {

        Path carpetaPdfs = Paths.get(rutaCarpetaPdfs).toAbsolutePath().normalize();
        Path carpetaTxt  = Paths.get(rutaCarpetaTxt).toAbsolutePath().normalize();

        // Validar carpeta origen
        if (!Files.exists(carpetaPdfs) || !Files.isDirectory(carpetaPdfs)) {
            System.out.println("[ERROR] Carpeta de PDFs inválida: " + carpetaPdfs);
            return;
        }

        // Crear carpeta destino si no existe
        try {
            Files.createDirectories(carpetaTxt);
            System.out.println("[INFO] Carpeta destino: " + carpetaTxt);
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo crear carpeta destino: " + e.getMessage());
            return;
        }

        // Buscar todos los PDFs en la carpeta origen
        List<Path> pdfs;
        try (Stream<Path> walk = Files.walk(carpetaPdfs)) {
            pdfs = walk
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString()
                            .toUpperCase(Locale.ROOT).endsWith(".PDF"))
                    .sorted()
                    .collect(Collectors.toList());
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo leer la carpeta: " + e.getMessage());
            return;
        }

        if (pdfs.isEmpty()) {
            System.out.println("[INFO] No se encontraron PDFs en: " + carpetaPdfs);
            return;
        }

        System.out.println("\n════════ EXTRACCIÓN PDF → TXT ════════");
        System.out.println("PDFs encontrados : " + pdfs.size());
        System.out.println("Carpeta destino  : " + carpetaTxt);

        // Configurar Tesseract
        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath(TESSDATA_PATH);
        tesseract.setLanguage(IDIOMA);
        tesseract.setPageSegMode(6);
        tesseract.setOcrEngineMode(1);

        int procesados = 0;
        int errores    = 0;

        for (Path pdf : pdfs) {
            System.out.println("\n[PROCESANDO] " + pdf.getFileName());
            try {
                // Paso 1: OCR → texto completo del PDF
                String textoCompleto = extraerTextoConOCR(pdf, tesseract);

                if (textoCompleto == null || textoCompleto.trim().isEmpty()) {
                    System.out.println("[SKIP] Sin texto reconocible: " + pdf.getFileName());
                    errores++;
                    continue;
                }

                // Paso 2: Determinar nombre del .txt destino
                // Mismo nombre que el PDF pero con extensión .txt
                String nombreTxt = pdf.getFileName().toString()
                        .replaceAll("(?i)\\.pdf$", ".txt");
                Path rutaTxt = carpetaTxt.resolve(nombreTxt);

                // Si ya existe, agregar sufijo para no sobreescribir
                if (Files.exists(rutaTxt)) {
                    String base = nombreTxt.replace(".txt", "");
                    nombreTxt   = base + "_" + System.currentTimeMillis() + ".txt";
                    rutaTxt     = carpetaTxt.resolve(nombreTxt);
                    System.out.println("[RENOMBRADO] Ya existe, guardando como: " + nombreTxt);
                }

                // Paso 3: Guardar el texto en el archivo .txt
                try (BufferedWriter writer = new BufferedWriter(
                        new OutputStreamWriter(
                                new FileOutputStream(rutaTxt.toFile()), "UTF-8"))) {
                    writer.write(textoCompleto);
                }

                System.out.println("[OK] Guardado: " + nombreTxt
                        + " (" + textoCompleto.length() + " caracteres)");
                procesados++;

            } catch (Exception e) {
                System.out.println("[ERROR] " + pdf.getFileName() + " → " + e.getMessage());
                errores++;
            }
        }

        System.out.println("\n════════ RESUMEN ════════");
        System.out.println("✅ PDFs procesados  : " + procesados);
        System.out.println("❌ Errores          : " + errores);
        System.out.println("📁 TXTs guardados en: " + carpetaTxt);
    }

    // ─────────────────────────────────────────────
    // NUEVO: PEGAR ARCHIVOS TXT DE CARPETA EN EXCEL
    // ─────────────────────────────────────────────

    /**
     * Lee todos los archivos .txt de una carpeta y pega su contenido
     * en un Excel destino: una fila por archivo, con columna "Archivo"
     * y columna "Contenido".
     *
     * Si el Excel destino ya existe, agrega filas al final.
     * Si no existe, lo crea con encabezados.
     *
     * Uso:
     *   ExtraccionCuadrillas ec = new ExtraccionCuadrillas();
     *   ec.pegarTxtEnExcel(
     *       "C:/mis_textos",
     *       "C:/resultado_textos.xlsx"
     *   );
     *
     * @param rutaCarpetaTxt   Carpeta con los archivos .txt
     * @param rutaExcelDestino Ruta del Excel donde se pegará el contenido
     */
    public void pegarTxtEnExcel(String rutaCarpetaTxt, String rutaExcelDestino) {

        Path carpetaTxt   = Paths.get(rutaCarpetaTxt).toAbsolutePath().normalize();
        Path archivoExcel = Paths.get(rutaExcelDestino).toAbsolutePath().normalize();

        // Validar carpeta origen
        if (!Files.exists(carpetaTxt) || !Files.isDirectory(carpetaTxt)) {
            System.out.println("[ERROR] Carpeta de TXTs inválida: " + carpetaTxt);
            return;
        }

        // Buscar todos los .txt en la carpeta
        List<Path> txts;
        try (Stream<Path> walk = Files.walk(carpetaTxt)) {
            txts = walk
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString()
                            .toUpperCase(Locale.ROOT).endsWith(".TXT"))
                    .sorted()
                    .collect(Collectors.toList());
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo leer la carpeta: " + e.getMessage());
            return;
        }

        if (txts.isEmpty()) {
            System.out.println("[INFO] No se encontraron archivos .txt en: " + carpetaTxt);
            return;
        }

        System.out.println("\n════════ PEGANDO TXTs EN EXCEL ════════");
        System.out.println("Archivos TXT encontrados : " + txts.size());
        System.out.println("Excel destino            : " + archivoExcel);

        // Cargar Excel existente o crear uno nuevo
        Workbook workbook;
        Sheet hoja;

        if (Files.exists(archivoExcel)) {
            try (FileInputStream fis = new FileInputStream(archivoExcel.toFile())) {
                workbook = new XSSFWorkbook(fis);
                hoja = workbook.getSheetAt(0);
                System.out.println("[INFO] Excel existente cargado. Filas actuales: " + hoja.getLastRowNum());
            } catch (Exception e) {
                System.out.println("[ERROR] No se pudo abrir el Excel: " + e.getMessage());
                return;
            }
        } else {
            workbook = new XSSFWorkbook();
            hoja = workbook.createSheet("Contenido TXT");

            // Encabezados
            CellStyle estilo = workbook.createCellStyle();
            Font fuente = workbook.createFont();
            fuente.setBold(true);
            estilo.setFont(fuente);
            estilo.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Row encabezado = hoja.createRow(0);
            String[] cols = {"Archivo", "Contenido"};
            for (int i = 0; i < cols.length; i++) {
                Cell celda = encabezado.createCell(i);
                celda.setCellValue(cols[i]);
                celda.setCellStyle(estilo);
            }
            hoja.setColumnWidth(0, 8000);
            hoja.setColumnWidth(1, 30000);
            System.out.println("[INFO] Excel nuevo creado con encabezados.");
        }

        int filaActual = hoja.getLastRowNum() + 1;
        int procesados = 0;
        int errores    = 0;

        // Procesar cada TXT
        for (Path txt : txts) {
            System.out.println("\n[PROCESANDO] " + txt.getFileName());
            try {
                // Leer todo el contenido del archivo
                StringBuilder contenido = new StringBuilder();
                try (BufferedReader reader = new BufferedReader(
                        new InputStreamReader(new FileInputStream(txt.toFile()), "UTF-8"))) {
                    String linea;
                    while ((linea = reader.readLine()) != null) {
                        contenido.append(linea).append("\n");
                    }
                }

                String textoFinal = contenido.toString().trim();

                if (textoFinal.isEmpty()) {
                    System.out.println("[SKIP] Archivo vacío: " + txt.getFileName());
                    errores++;
                    continue;
                }

                // Escribir fila: columna A = nombre archivo, columna B = contenido
                Row fila = hoja.createRow(filaActual++);
                fila.createCell(0).setCellValue(txt.getFileName().toString());
                fila.createCell(1).setCellValue(textoFinal);

                // Alto de fila proporcional al contenido (máx 400 puntos)
                int lineas = textoFinal.split("\n").length;
                float alto = Math.min(lineas * 15f, 400f);
                fila.setHeightInPoints(alto);

                System.out.println("[OK] " + txt.getFileName() + " (" + lineas + " líneas)");
                procesados++;

            } catch (Exception e) {
                System.out.println("[ERROR] " + txt.getFileName() + " → " + e.getMessage());
                errores++;
            }
        }

        // Guardar el Excel
        try (FileOutputStream fos = new FileOutputStream(archivoExcel.toFile())) {
            workbook.write(fos);
            workbook.close();
            System.out.println("\n════════ RESUMEN ════════");
            System.out.println("✅ Archivos pegados : " + procesados);
            System.out.println("❌ Errores          : " + errores);
            System.out.println("💾 Excel guardado   : " + archivoExcel);
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo guardar el Excel: " + e.getMessage());
        }
    }

    // ─────────────────────────────────────────────
    // NUEVO: CREAR ENCABEZADOS EN EXCEL NUEVO
    // ─────────────────────────────────────────────

    private void crearEncabezadosExcel(Sheet hoja, Workbook workbook) {
        String[] encabezados = {
                "Archivo PDF",
                "#", "Descripcion", "Codigo", "Almacen", "Medida",
                "Precio Unitario", "Cantidad", "Precio Total", "Proveedor",
                "Cuenta", "Proyecto", "Centro Costo 1", "Centro Costo 2",
                "Centro Costo 3", "Centro Costo 4", "Centro Costo 5", "Referencia"
        };

        CellStyle estilo = workbook.createCellStyle();
        Font fuente = workbook.createFont();
        fuente.setBold(true);
        estilo.setFont(fuente);
        estilo.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row fila = hoja.createRow(0);
        for (int i = 0; i < encabezados.length; i++) {
            Cell celda = fila.createCell(i);
            celda.setCellValue(encabezados[i]);
            celda.setCellStyle(estilo);
            hoja.setColumnWidth(i, 4500);
        }
    }
}
