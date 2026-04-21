package nombradas;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Agrupa archivos “NMBRADA” y crea un resumen y hojas detalladas por nave.
 * Reglas:
 *  - Solo procesa la ULTIMA versión por (nave, fecha, jornada).
 *  - Resumen agrupa por (nave, fecha) y suma personal de las jornadas (ya depuradas por versión).
 *  - Hojas por nave muestran bloques por fecha, con los trabajadores provenientes de las jornadas (última versión).
 */
public class AgruparNombradas {

    private String rutaCarpeta;
    private String rutaSalida;

    public AgruparNombradas(String rutaCarpeta, String rutaSalida) {
        this.rutaCarpeta = rutaCarpeta;
        this.rutaSalida = rutaSalida;
    }

    // ==========================================================
    // LECTURA PRINCIPAL
    // ==========================================================
    public void leerExcels() {
        try {
            // Java 8: usar Paths.get en lugar de Path.of
            Path dir = Paths.get(rutaCarpeta).toAbsolutePath().normalize();
            if (!Files.exists(dir) || !Files.isDirectory(dir)) {
                System.out.println("[ERROR] Ruta inválida: " + dir);
                return;
            }

            List<Path> excels;
            try (Stream<Path> walk = Files.walk(dir)) {
                excels = walk
                        .filter(Files::isRegularFile)
                        .filter(this::esExcelNombrada)
                        .sorted()
                        .collect(Collectors.toList());
            }

            if (excels.isEmpty()) {
                System.out.println("[INFO] No se encontraron archivos NMBRADA");
                return;
            }

            // ----------------------------------------------------------------
            // 1) ELEGIR SOLO LA ULTIMA VERSIÓN por (nave | fecha | jornada)
            //    key = nave|fecha|jor
            // ----------------------------------------------------------------
            Map<String, RegistroSeleccion> seleccion = new HashMap<String, RegistroSeleccion>();

            for (Path archivo : excels) {
                InfoNombrada info;
                try {
                    info = parsearNombreArchivo(archivo);
                } catch (Exception ex) {
                    System.out.println("[WARN] Saltando archivo (parse falló): " + archivo.getFileName() + " -> " + ex.getMessage());
                    continue;
                }

                String clave = info.nombreNave + "|" + info.fecha + "|JOR" + info.jornada;
                seleccion.compute(clave, (k, v) -> {
                    if (v == null) return new RegistroSeleccion(archivo, info);
                    // Mantener el de mayor versión
                    if (info.version > v.info.version) return new RegistroSeleccion(archivo, info);
                    return v;
                });
            }

            if (seleccion.isEmpty()) {
                System.out.println("[INFO] No hay archivos válidos tras aplicar selección por versión.");
                return;
            }

            // ----------------------------------------------------------------
            // 2) Construir RESUMEN (nave|fecha) y DETALLE (nave -> fecha -> lista trabajadores)
            //    usando SOLO los archivos seleccionados (última versión)
            // ----------------------------------------------------------------
            Map<String, ResultadoNombrada> resumen = new HashMap<String, ResultadoNombrada>();
            Map<String, Map<String, List<Trabajador>>> detalle = new HashMap<String, Map<String, List<Trabajador>>>();

            for (RegistroSeleccion rs : seleccion.values()) {
                Path archivo = rs.path;
                InfoNombrada info = rs.info;

                List<Trabajador> trabajadores = extraerTrabajadores(archivo);

                // Resumen: acumular por nave|fecha
                String kResumen = info.nombreNave + "|" + info.fecha;
                resumen.compute(kResumen, (k, v) -> {
                    int cant = trabajadores.size();
                    if (v == null) return new ResultadoNombrada(info.nombreNave, info.fecha, cant);
                    v.personal += cant;
                    return v;
                });

                // Detalle: nave -> fecha -> lista (concatenamos jornadas de esa fecha)
                if (!detalle.containsKey(info.nombreNave)) {
                    detalle.put(info.nombreNave, new HashMap<String, List<Trabajador>>());
                }
                Map<String, List<Trabajador>> porFecha = detalle.get(info.nombreNave);
                if (!porFecha.containsKey(info.fecha)) {
                    porFecha.put(info.fecha, new ArrayList<Trabajador>());
                }
                porFecha.get(info.fecha).addAll(trabajadores);
            }

            // ----------------------------------------------------------------
            // 3) Generar archivo final (Resumen + hojas por nave)
            // ----------------------------------------------------------------
            generarArchivoResumenYDetalle(new ArrayList<ResultadoNombrada>(resumen.values()), detalle);

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // ==========================================================
    // VALIDACIÓN DE ARCHIVOS
    // ==========================================================
    private boolean esExcelNombrada(Path p) {
        String f = p.getFileName().toString().toUpperCase(Locale.ROOT);
        if (f.startsWith("~$")) return false;
        if (!f.startsWith("NOMB")) return false;
        return f.endsWith(".XLS") || f.endsWith(".XLSX");
    }

    // ==========================================================
    // PARSEO DEL NOMBRE DEL ARCHIVO
    //  - Fecha: dd.MM.yy en cualquier posición
    //  - Versión: token justo tras "NMBRADA", con forma ####-N → versión = N
    //  - Jornada: "JOR <n>"
    //  - Nave: tokens entre el id y la fecha
    // ==========================================================
    private InfoNombrada parsearNombreArchivo(Path p) {
        String original = p.getFileName().toString();
        String name = original.replace(".xlsx", "").replace(".xls", "").trim();

        // Normalizar múltiples espacios
        String[] parts = name.split("\\s+");

        // 1) Validar y extraer versión desde el segundo token (####-N)
        //    Estructura esperada: NMBRADA <id-version> <nave ...> <fecha> JOR <n> ...
        if (parts.length < 3 || !parts[0].equalsIgnoreCase("NOMBRADA"))
            throw new IllegalArgumentException("Nombre no inicia con 'NMBRADA': " + original);

        int version = 0;
        String idToken = parts[1]; // esperado "0001-2" o similar
        Matcher mVer = Pattern.compile("^(\\d+)-(\\d+)$").matcher(idToken);
        if (mVer.find()) {
            // String call = mVer.group(1); // no usado
            version = Integer.parseInt(mVer.group(2));
        } else {
            // Si no hay patrón estricto, dejamos version=0 (entra pero no ganará frente a válidos)
            System.out.println("[WARN] No se pudo extraer versión de: " + original);
        }

        // 2) Fecha (posición variable)
        Matcher mFecha = Pattern.compile("(\\d{2}\\.\\d{2}\\.\\d{2})").matcher(name);
        if (!mFecha.find())
            throw new IllegalArgumentException("No se encontró fecha dd.MM.yy en: " + original);
        String fecha = mFecha.group(1);

        // 3) Jornada (buscar "JOR n")
        Matcher mJor = Pattern.compile("JOR\\s+(\\d+)", Pattern.CASE_INSENSITIVE).matcher(name);
        int jornada = 0; // 0 si no se encuentra (evita NPE y sigue agrupando)
        if (mJor.find()) {
            try { jornada = Integer.parseInt(mJor.group(1)); } catch (NumberFormatException ignored) {}
        } else {
            System.out.println("[WARN] No se detectó 'JOR n' en: " + original + " → se usará JOR0");
        }

        // 4) Índice de la fecha entre tokens
        int idxFecha = -1;
        for (int i = 0; i < parts.length; i++) {
            if (parts[i].contains(fecha)) { idxFecha = i; break; }
        }
        if (idxFecha < 0) throw new IllegalArgumentException("No se pudo posicionar la fecha en tokens: " + original);

        // 5) Nave: tokens desde índice 2 (después de NMBRADA y ####-N) hasta antes de la fecha
        StringBuilder nave = new StringBuilder();
        for (int i = 2; i < idxFecha; i++) nave.append(parts[i]).append(" ");
        String nombreNave = nave.toString().trim();

        return new InfoNombrada(nombreNave, fecha, jornada, version);
    }

    // ==========================================================
    // EXTRAER TRABAJADORES desde la fila 52 (índice 51)
    //  Copiando VALORES (no fórmulas) con DataFormatter + FormulaEvaluator
    //  Columnas: 0,1,2,3,5,6
    // ==========================================================
    private List<Trabajador> extraerTrabajadores(Path archivo) {
        List<Trabajador> lista = new ArrayList<Trabajador>();
        try (FileInputStream fis = new FileInputStream(archivo.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // SIN DataFormatter, SIN FormulaEvaluator → solo valores cacheados
            Sheet sheet = wb.getSheetAt(0);
            int inicio = 52 - 1;
            int last   = sheet.getLastRowNum();

            for (int i = inicio; i <= last; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String nro   = getCellValue(row.getCell(0));
                String apePa = getCellValue(row.getCell(1));
                String apeMa = getCellValue(row.getCell(2));
                String priNo = getCellValue(row.getCell(3));
                String dni   = getCellValue(row.getCell(5));
                String cargo = getCellValue(row.getCell(6));

                if ((nro + apePa + apeMa + priNo + dni + cargo).isEmpty()) continue;

                lista.add(new Trabajador(nro, apePa, apeMa, priNo, dni, cargo));
            }

        } catch (Exception e) {
            System.out.println("[ERROR] Leyendo " + archivo.getFileName() + ": " + e.getMessage());
        }
        return lista;
    }

    // ==========================================================
    // CONSTRUCCIÓN DEL ARCHIVO: Resumen + Hojas por Nave
    // ==========================================================
    private void generarArchivoResumenYDetalle(List<ResultadoNombrada> resumen,
                                               Map<String, Map<String, List<Trabajador>>> detalle) {
        try {
            // Java 8: Paths.get en lugar de Path.of
            Path dir = Paths.get(rutaSalida);
            if (!Files.exists(dir)) Files.createDirectories(dir);
            Path out = dir.resolve("Resumen Nombradas.xlsx");

            XSSFWorkbook wb = new XSSFWorkbook();

            // 1) Hoja Resumen
            generarHojaResumen(wb, resumen);

            // 2) Hojas por nave
            generarHojasPorNave(wb, detalle);

            // Guardar
            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(out.toFile())) {
                wb.write(fos);
            }
            wb.close();

            System.out.println("[OK] Archivo generado con Resumen + hojas por nave: " + out);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case FORMULA:
                // Leer el valor cacheado (ya calculado por Excel), NO re-evaluar
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:  return cell.getRichStringCellValue().getString().trim();
                    case NUMERIC:
                        // Evitar notación científica en números como DNI
                        double d = cell.getNumericCellValue();
                        if (d == Math.floor(d)) return String.valueOf((long) d);
                        return String.valueOf(d);
                    case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
                    case ERROR:   return ""; // celda con error → ignorar
                    default:      return "";
                }
            case STRING:  return cell.getRichStringCellValue().getString().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d)) return String.valueOf((long) d);
                return String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
            default:      return "";
        }
    }

    // ----------------------------------------------------------
    // Hoja "Resumen"
    // ----------------------------------------------------------
    private void generarHojaResumen(XSSFWorkbook wb, List<ResultadoNombrada> lista) {
        Sheet sheet = wb.createSheet("Resumen");

        CellStyle head = crearEstiloEncabezadoAzul(wb);

        Row h = sheet.createRow(0);
        String[] cols = {"Nave", "Fecha", "Personal Nombrado"};
        for (int i = 0; i < cols.length; i++) {
            Cell c = h.createCell(i);
            c.setCellValue(cols[i]);
            c.setCellStyle(head);
        }

        // Orden por nave y fecha (yy-MM-dd)
        Collections.sort(lista, Comparator
                .comparing((ResultadoNombrada r) -> r.nave)
                .thenComparing(new java.util.function.Function<ResultadoNombrada, String>() {
                    @Override public String apply(ResultadoNombrada r) { return normalizarFechaOrden(r.fecha); }
                }));

        int r = 1;
        for (ResultadoNombrada x : lista) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(x.nave);
            row.createCell(1).setCellValue(x.fecha);
            row.createCell(2).setCellValue(x.personal);
        }

        for (int i = 0; i < cols.length; i++) sheet.autoSizeColumn(i);
        sheet.createFreezePane(0, 1);
    }

    // ----------------------------------------------------------
    // Hojas por nave con bloques por FECHA
    // ----------------------------------------------------------
    private void generarHojasPorNave(XSSFWorkbook wb,
                                     Map<String, Map<String, List<Trabajador>>> detalle) {

        CellStyle head = crearEstiloEncabezadoAzul(wb);
        CellStyle dateStyle = crearEstiloFechaSuave(wb);

        List<String> naves = new ArrayList<String>(detalle.keySet());
        Collections.sort(naves);

        for (String nave : naves) {
            String sheetName = nombreHojaSeguro(wb, nave);
            Sheet sheet = wb.createSheet(sheetName);

            Row h = sheet.createRow(0);
            String[] cols = {"N° operador", "Ape. Paterno", "Ape. Materno", "Pri. Nombre", "DNI", "Cargo"};
            for (int i = 0; i < cols.length; i++) {
                Cell c = h.createCell(i);
                c.setCellValue(cols[i]);
                c.setCellStyle(head);
            }

            Map<String, List<Trabajador>> porFecha = detalle.get(nave);
            List<String> fechas = new ArrayList<String>(porFecha.keySet());
            Collections.sort(fechas, new Comparator<String>() {
                @Override public int compare(String a, String b) {
                    return normalizarFechaOrden(a).compareTo(normalizarFechaOrden(b));
                }
            });

            int rowIdx = 1;

            for (String fecha : fechas) {
                // Fila combinada de FECHA
                Row rf = sheet.createRow(rowIdx);
                Cell cf = rf.createCell(0);
                cf.setCellValue("Fecha: " + fecha);
                cf.setCellStyle(dateStyle);
                sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 5));
                rowIdx++;

                // Detalle
                List<Trabajador> trabajadores = porFecha.get(fecha);
                for (Trabajador t : trabajadores) {
                    Row rw = sheet.createRow(rowIdx++);
                    rw.createCell(0).setCellValue(t.nroOperador);
                    rw.createCell(1).setCellValue(t.apePaterno);
                    rw.createCell(2).setCellValue(t.apeMaterno);
                    rw.createCell(3).setCellValue(t.priNombre);
                    rw.createCell(4).setCellValue(t.dni);
                    rw.createCell(5).setCellValue(t.cargo);
                }
            }

            for (int i = 0; i < 6; i++) sheet.autoSizeColumn(i);
            sheet.createFreezePane(0, 1);
        }
    }

    // ==========================================================
    // Utilidades
    // ==========================================================
    private String normalizarFechaOrden(String ddmmyy) {
        try {
            String[] p = ddmmyy.split("\\.");
            String dd = p[0], mm = p[1], yy = p[2];
            return yy + "-" + mm + "-" + dd;
        } catch (Exception e) {
            return ddmmyy;
        }
    }

    private String nombreHojaSeguro(Workbook wb, String base) {
        String clean = base.replaceAll("[\\\\/?*\\[\\]:]", " ").trim();
        clean = clean.length() > 31 ? clean.substring(0, 31) : clean;

        String name = clean;
        int suf = 2;
        while (wb.getSheet(name) != null) {
            String suffix = " (" + suf++ + ")";
            int max = 31 - suffix.length();
            String cut = clean.length() > max ? clean.substring(0, max) : clean;
            name = cut + suffix;
        }
        return name;
    }

    private CellStyle crearEstiloEncabezadoAzul(XSSFWorkbook wb) {
        Font f = wb.createFont();
        f.setBold(true);
        f.setColor(IndexedColors.WHITE.getIndex());

        CellStyle s = wb.createCellStyle();
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private CellStyle crearEstiloFechaSuave(XSSFWorkbook wb) {
        Font f = wb.createFont();
        f.setBold(false);
        f.setColor(IndexedColors.BLACK.getIndex());

        CellStyle s = wb.createCellStyle();
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    // ==========================================================
    // POJOs / registros
    // ==========================================================
    private static class InfoNombrada {
        String nombreNave;
        String fecha;
        int jornada;
        int version;
        InfoNombrada(String n, String f, int j, int v) {
            this.nombreNave = n; this.fecha = f; this.jornada = j; this.version = v;
        }
    }

    private static class ResultadoNombrada {
        String nave;
        String fecha;
        int personal;
        ResultadoNombrada(String n, String f, int p) { this.nave = n; this.fecha = f; this.personal = p; }
    }

    private static class Trabajador {
        String nroOperador;
        String apePaterno;
        String apeMaterno;
        String priNombre;
        String dni;
        String cargo;
        Trabajador(String nro, String ap, String am, String pn, String dni, String cg) {
            this.nroOperador = nro; this.apePaterno = ap; this.apeMaterno = am;
            this.priNombre = pn; this.dni = dni; this.cargo = cg;
        }
    }

    private static class RegistroSeleccion {
        Path path;
        InfoNombrada info;
        RegistroSeleccion(Path p, InfoNombrada i) { this.path = p; this.info = i; }
    }
}
