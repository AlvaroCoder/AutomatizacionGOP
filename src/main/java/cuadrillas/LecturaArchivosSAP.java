package cuadrillas;

import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripper;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class LecturaArchivosSAP {
    public static class ResultadoOCR {
        public final String rutaArchivo;
        public final boolean tieneTexto;
        public final boolean tieneImagenes;
        public final int totalPaginas;
        public final int paginasConTexto;
        public final String estado; // "LEGIBLE", "SOLO_IMAGEN", "MIXTO", "VACIO"

        public ResultadoOCR(String ruta, boolean tieneTexto, boolean tieneImagenes,
                            int totalPaginas, int paginasConTexto) {
            this.rutaArchivo = ruta;
            this.tieneTexto = tieneTexto;
            this.tieneImagenes = tieneImagenes;
            this.totalPaginas = totalPaginas;
            this.paginasConTexto = paginasConTexto;

            if (tieneTexto && tieneImagenes) {
                this.estado = "MIXTO";
            } else if (tieneTexto) {
                this.estado = "LEGIBLE";
            } else if (tieneImagenes) {
                this.estado = "SOLO_IMAGEN (necesita OCR)";
            } else {
                this.estado = "VACIO";
            }
        }

        @Override
        public String toString() {
            return String.format(
                    "[%s] Páginas: %d | Con texto: %d | Imágenes: %s | Estado: %s",
                    rutaArchivo, totalPaginas, paginasConTexto,
                    tieneImagenes ? "Sí" : "No", estado
            );
        }
    }

    private String rutaCarpetaNaves = "";

    public LecturaArchivosSAP() {}

    public ResultadoOCR validarOCR(Path rutaPdf) {
        final int UMBRAL_CARACTERES = 20; // mínimo de chars para considerar que hay texto real

        try (PDDocument doc = PDDocument.load(rutaPdf.toFile())) {
            int totalPaginas = doc.getNumberOfPages();
            int paginasConTexto = 0;
            boolean tieneImagenes = false;

            PDFTextStripper stripper = new PDFTextStripper();

            for (int i = 1; i <= totalPaginas; i++) {
                // Extraer texto de esta página
                stripper.setStartPage(i);
                stripper.setEndPage(i);
                String textoPagina = stripper.getText(doc).trim();

                if (textoPagina.length() >= UMBRAL_CARACTERES) {
                    paginasConTexto++;
                }

                // Detectar si la página tiene imágenes embebidas
                PDPage page = doc.getPage(i - 1);
                if (paginaTieneImagenes(page)) {
                    tieneImagenes = true;
                }
            }

            boolean tieneTexto = paginasConTexto > 0;
            return new ResultadoOCR(
                    rutaPdf.getFileName().toString(),
                    tieneTexto, tieneImagenes,
                    totalPaginas, paginasConTexto
            );

        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo analizar: " + rutaPdf.getFileName() + " → " + e.getMessage());
            return null;
        }
    }

    private boolean paginaTieneImagenes(PDPage page) {
        try {
            // Revisamos los recursos de la página buscando XObjects de tipo imagen
            if (page.getResources() == null) return false;
            for (COSName name : page.getResources().getXObjectNames()) {
                if (page.getResources().isImageXObject(name)) {
                    return true;
                }
            }
        } catch (Exception e) {
            System.out.println("e = " + e.getMessage());
            // Si no se puede leer, asumimos que no hay imágenes
        }
        return false;
    }

    public void copiarPdfsADestino(String rutaDestino) {
        try {
            Path carpetaOrigen = Paths.get(this.rutaCarpetaNaves).toAbsolutePath().normalize();
            Path carpetaDestino = Paths.get(rutaDestino).toAbsolutePath().normalize();

            // Validar carpeta origen
            if (!Files.exists(carpetaOrigen) || !Files.isDirectory(carpetaOrigen)) {
                System.out.println("[ERROR] Ruta de origen inválida: " + carpetaOrigen);
                return;
            }

            // Crear carpeta destino si no existe
            if (!Files.exists(carpetaDestino)) {
                Files.createDirectories(carpetaDestino);
                System.out.println("[INFO] Carpeta destino creada: " + carpetaDestino);
            }

            String patronEsperado = "SAP";

            // Buscar PDFs en todas las subcarpetas
            List<Path> archivosEncontrados;
            try (Stream<Path> walk = Files.walk(carpetaOrigen)) {
                archivosEncontrados = walk
                        .filter(Files::isRegularFile)
                        .filter(path -> {
                            String nombre = path.getFileName().toString().toUpperCase(Locale.ROOT);
                            return nombre.startsWith(patronEsperado) && nombre.endsWith(".PDF");
                        })
                        .collect(Collectors.toList());
            }

            if (archivosEncontrados.isEmpty()) {
                System.out.println("[INFO] No se encontraron PDFs con el patrón '" + patronEsperado + "'");
                return;
            }

            System.out.println("\n════════ COPIANDO PDFs ════════");
            System.out.println("Origen  : " + carpetaOrigen);
            System.out.println("Destino : " + carpetaDestino);
            System.out.println("Total   : " + archivosEncontrados.size() + " archivos\n");

            int copiados = 0;
            int omitidos = 0;
            int errores  = 0;

            for (Path archivoOrigen : archivosEncontrados) {
                try {
                    String nombreArchivo = archivoOrigen.getFileName().toString();
                    Path archivoDestino  = carpetaDestino.resolve(nombreArchivo);

                    // Si ya existe un archivo con el mismo nombre → renombrar con sufijo
                    if (Files.exists(archivoDestino)) {
                        String nombreSinExt = nombreArchivo.replaceAll("(?i)\\.pdf$", "");
                        String nuevoNombre  = nombreSinExt + "_copia_" + System.currentTimeMillis() + ".pdf";
                        archivoDestino      = carpetaDestino.resolve(nuevoNombre);
                        System.out.println("[RENOMBRADO] " + nombreArchivo + " → " + nuevoNombre);
                        omitidos++;
                    }

                    Files.copy(archivoOrigen, archivoDestino, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("[OK] " + nombreArchivo);
                    copiados++;

                } catch (Exception e) {
                    System.out.println("[ERROR] No se pudo copiar: " + archivoOrigen.getFileName() + " → " + e.getMessage());
                    errores++;
                }
            }

            // Resumen final
            System.out.println("\n════════ RESUMEN DE COPIA ════════");
            System.out.println("✅ Copiados correctamente : " + copiados);
            System.out.println("⚠️  Renombrados (duplicados): " + omitidos);
            System.out.println("❌ Errores                : " + errores);

        } catch (Exception err) {
            System.out.println("[ERROR] " + err.getMessage());
        }
    }

    public void listarDocumentosPdf() {
        try {
            Path carpetaNaves = Paths.get(this.rutaCarpetaNaves).toAbsolutePath().normalize();
            if (!Files.exists(carpetaNaves) || !Files.isDirectory(carpetaNaves)) {
                System.out.println("[ERROR] Ruta inválida: " + carpetaNaves);
                return;
            }

            String patronEsperado = "SAP";

            List<Path> archivosNaves;
            try (Stream<Path> walk = Files.walk(carpetaNaves)) {
                archivosNaves = walk
                        .filter(Files::isRegularFile)
                        .filter(path -> {
                            String nombre = path.getFileName().toString().toUpperCase(Locale.ROOT);
                            return nombre.startsWith(patronEsperado) && nombre.endsWith(".PDF");
                        })
                        .collect(Collectors.toList());
            }

            if (archivosNaves.isEmpty()) {
                System.out.println("No se encontraron archivos PDF que empiecen con '" + patronEsperado + "'");
                return;
            }

            System.out.println("Archivos encontrados: " + archivosNaves.size());
            System.out.println("\n════════ VALIDACIÓN OCR ════════");

            // Clasificar PDFs
            List<Path> legibles = new ArrayList<>();
            List<Path> soloImagen = new ArrayList<>();

            for (Path pdf : archivosNaves) {
                ResultadoOCR resultado = validarOCR(pdf);
                if (resultado != null) {
                    System.out.println(resultado);
                    if (resultado.estado.startsWith("LEGIBLE") || resultado.estado.equals("MIXTO")) {
                        legibles.add(pdf);
                    } else {
                        soloImagen.add(pdf);
                    }
                }
            }

            // Resumen final
            System.out.println("\n════════ RESUMEN ════════");
            System.out.println("✅ PDFs legibles (con texto):     " + legibles.size());
            System.out.println("⚠️  PDFs solo imagen (sin OCR):   " + soloImagen.size());

            if (!soloImagen.isEmpty()) {
                System.out.println("\nArchivos que necesitan OCR:");
                soloImagen.forEach(p -> System.out.println("   - " + p.getFileName()));
            }

            // Demo: leer el primer PDF legible
            if (!legibles.isEmpty()) {
                System.out.println("\n════════ LECTURA DE PRUEBA ════════");
                System.out.println("Leyendo: " + legibles.get(0).getFileName());
                leerPdf(legibles.get(0));
            }

        } catch (Exception err) {
            System.out.println("[ERROR] " + err.getMessage());
        }
    }

    private void leerPdf(Path rutaPdf) {
        try (PDDocument documento = PDDocument.load(rutaPdf.toFile())) {
            PDFTextStripper stripper = new PDFTextStripper();
            String texto = stripper.getText(documento);

            String[] lineas = texto.split("\r\n|\r|\n");
            for (int i = 0; i < Math.min(lineas.length, 10); i++) {
                System.out.println(lineas[i]);
            }
            if (lineas.length > 10) {
                System.out.println("... (contenido truncado, " + lineas.length + " líneas totales)");
            }

        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo leer el PDF: " + e.getMessage());
        }
    }



    public LecturaArchivosSAP(String rutaCarpetaNaves) {
        this.rutaCarpetaNaves = rutaCarpetaNaves;
    }

    public void setRutaCarpetaNaves(String rutaCarpetaNaves) {
        this.rutaCarpetaNaves = rutaCarpetaNaves;
    }
}

