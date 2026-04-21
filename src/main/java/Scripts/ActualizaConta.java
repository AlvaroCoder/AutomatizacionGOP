package Scripts;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class ActualizaConta {
    private String RutaThroughPut;
    private String RutaObjetivo;
    private int ColumnaSeleccionada;
    private int SheetThroughPut = 0;
    private int SheetObjetivo = 22;
    private int colInicioObjetivo = 45;
    private String objetivoAnchorRef;  // p.ej. "CELDA_BASE_CONTA" o "HojaDestino!AT8"

    private XSSFSheet ThroughPutSheet, ObjetivoSheet;

    public ActualizaConta(String RutaThrough, String RutaObjetivo){
        this.RutaObjetivo = RutaObjetivo;
        this.RutaThroughPut = RutaThrough;
    }

    public ActualizaConta(String RutaThrough, String RutaObjetivo, int ColumnaSeleccionada){
        this.RutaObjetivo = RutaObjetivo;
        this.RutaThroughPut = RutaThrough;
        this.ColumnaSeleccionada = ColumnaSeleccionada;
    }
    public void setColumnaSeleccionada(int columnaSeleccionada) {
        ColumnaSeleccionada = columnaSeleccionada;
    }

    public void setSheetThroughPut(int sheetThroughPut) {
        SheetThroughPut = sheetThroughPut;
    }

    // (Opcional) mantener compatibilidad hacia atrás:
    @Deprecated
    public void setSheetObjetivo(int sheetObjetivo) {
        // Mantengo por compatibilidad, pero lo ideal es usar setObjetivoAnchor(...)
        SheetObjetivo = sheetObjetivo;
    }

    // NUEVO: establece el ancla (nombre definido o A1 con hoja)
    public void setObjetivoAnchor(String anchorRef) {
        this.objetivoAnchorRef = anchorRef;
    }


    public void setColInicioObjetivo(int colInicioObjetivo) {
        this.colInicioObjetivo = colInicioObjetivo;
    }

    public void ejecutarRango(int filaInicio, int filaFin) {

        final int COL_VESSEL_OBJ = 3;
        final int ROW_HEADER_THROUGH = 8;
        final int ROW_DATA_START = 30;
        final int ROW_DATA_END_EXCLUSIVE = 39; // 30..38 inclusive

        try (
                FileInputStream throughStream = new FileInputStream(this.RutaThroughPut);
                XSSFWorkbook throughWB = new XSSFWorkbook(throughStream);
                FileInputStream objetivoStream = new FileInputStream(this.RutaObjetivo);
                XSSFWorkbook objetivoWB = new XSSFWorkbook(objetivoStream)
        ) {

            DataFormatter formatter = new DataFormatter();

            // Usa el sheet configurado (si te interesa cambiarlo, llama a setSheetThroughPut)
            XSSFSheet sheetThrough = throughWB.getSheetAt(this.SheetThroughPut);

            // 🔹 Resolver Hoja Objetivo + columna de inicio y especial a partir del ancla
            ResolvedAnchor anchor = resolveAnchor(objetivoWB);
            XSSFSheet sheetObjetivo = anchor.sheet;
            final int COL_DESTINO_START = anchor.startCol;
            final int COL_DESTINO_ESPECIAL = anchor.specialCol;

            Row header = sheetThrough.getRow(ROW_HEADER_THROUGH);
            if (header == null) {
                throw new IllegalStateException("Header ThroughPut no existe");
            }

            Map<String, Integer> headerMap = new HashMap<>();
            for (Cell c : header) {
                String key = formatter.formatCellValue(c).trim();
                if (!key.isEmpty()) {
                    headerMap.put(key, c.getColumnIndex());
                    headerMap.put("0" + key, c.getColumnIndex());
                }
            }

            System.out.println("Se inició el proceso");

            for (int i = filaInicio - 1; i < filaFin; i++) {

                Row vesselRow = sheetObjetivo.getRow(i);
                if (vesselRow == null) continue;

                Cell vesselCell = vesselRow.getCell(COL_VESSEL_OBJ);
                if (vesselCell == null) continue;

                String vesselName = formatter.formatCellValue(vesselCell).trim();
                if (vesselName.isEmpty()) continue;
                System.out.println("vesselName = " + vesselName);

                Integer colMatch = headerMap.get(vesselName);
                if (colMatch == null) continue;

                int colDestino = COL_DESTINO_START;

                // Recorremos SIEMPRE 30..38 y escribimos; si no hay dato => en blanco
                for (int j = ROW_DATA_START; j < ROW_DATA_END_EXCLUSIVE; j++) {

                    int colFinal;
                    if (j == 35) {           // tu salto +2 columnas
                        colFinal = colDestino;
                        colDestino += 3;
                    } else if (j == 38) {    // columna especial relativa
                        colFinal = COL_DESTINO_ESPECIAL;
                    } else {
                        colFinal = colDestino++;
                    }

                    Cell destino = vesselRow.getCell(colFinal);
                    if (destino == null) destino = vesselRow.createCell(colFinal);

                    // Obtener origen
                    Row rowData = sheetThrough.getRow(j);
                    Cell origen = (rowData == null ? null : rowData.getCell(colMatch));

                    String asText = (origen == null) ? "" : formatter.formatCellValue(origen).trim();

                    if (asText.isEmpty()) {
                        // 🔸 sin data => celda en blanco (no "0")
                        destino.setCellValue(0);  // compatible con POI antiguos
                    } else {
                        // Intentar número; si no, dejar texto
                        try {
                            double val = Double.parseDouble(asText.replace(",", ""));
                            destino.setCellValue(val);
                        } catch (NumberFormatException nfe) {
                            // No es número, escribimos el texto tal cual (o déjalo en blanco si prefieres)
                            destino.setCellValue(asText);
                        }
                    }
                }

                System.out.println("Datos copiados para fila " + (i + 1));
            }

            objetivoWB.setForceFormulaRecalculation(true);

            try (FileOutputStream fos = new FileOutputStream(RutaObjetivo)) {
                objetivoWB.write(fos);
            }

            System.out.println("Proceso finalizado correctamente");

        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
        }
    }

    // ---- Helper interno para resolver hoja y columna a partir del ancla ----
    private static class ResolvedAnchor {
        final XSSFSheet sheet;
        final int startCol;       // Columna donde empieza a escribir (antes colInicioObjetivo)
        final int specialCol;     // Columna "especial" relativa (antes fija en 55)

        ResolvedAnchor(XSSFSheet sheet, int startCol, int specialCol) {
            this.sheet = sheet;
            this.startCol = startCol;
            this.specialCol = specialCol;
        }
    }

    private ResolvedAnchor resolveAnchor(XSSFWorkbook objetivoWB) {
        // Si no se indicó ancla, uso el comportamiento anterior
        if (objetivoAnchorRef == null || objetivoAnchorRef.trim().isEmpty()) {
            XSSFSheet sh = objetivoWB.getSheetAt(this.SheetObjetivo);
            // Mantener compatibilidad: especial = inicio + (55 - 45) = +10
            int special = this.colInicioObjetivo + 10;
            return new ResolvedAnchor(sh, this.colInicioObjetivo, special);
        }

        // 1) Intentar como "Named Range"
        System.out.println("objetivoAnchorRef = " + objetivoAnchorRef);
        try {
            org.apache.poi.ss.usermodel.Name nm = objetivoWB.getName(objetivoAnchorRef);
            if (nm != null) {
                String ref = nm.getRefersToFormula();              // p.ej. 'HojaX'!$AT$8
                org.apache.poi.ss.util.AreaReference ar =
                        new org.apache.poi.ss.util.AreaReference(ref, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
                org.apache.poi.ss.util.CellReference first = ar.getFirstCell();
                String sheetName = first.getSheetName();
                XSSFSheet sh = objetivoWB.getSheet(sheetName);
                int start = first.getCol(); // 0-based
                // Offset relativo que antes era fijo (55 - 45) = 10
                int special = start + 10;
                return new ResolvedAnchor(sh, start, special);
            }
        } catch (Exception ignore) { /* caerá al parseo A1 */ }

        // 2) Parsear como referencia A1: "HojaDestino!AT8" (recomendado)
        try {
            org.apache.poi.ss.util.CellReference cr = new org.apache.poi.ss.util.CellReference(objetivoAnchorRef);
            String sheetName = cr.getSheetName();
            if (sheetName == null) {
                // Si no viene hoja en la ref, caigo al índice existente por compatibilidad
                XSSFSheet sh = objetivoWB.getSheetAt(this.SheetObjetivo);
                int start = cr.getCol();
                int special = start + 10;
                return new ResolvedAnchor(sh, start, special);
            } else {
                XSSFSheet sh = objetivoWB.getSheet(sheetName);
                int start = cr.getCol();
                int special = start + 10;
                return new ResolvedAnchor(sh, start, special);
            }
        } catch (Exception e) {
            // Fallback total: comportamiento anterior
            XSSFSheet sh = objetivoWB.getSheetAt(this.SheetObjetivo);
            int special = this.colInicioObjetivo + 10;
            return new ResolvedAnchor(sh, this.colInicioObjetivo, special);
        }
    }

    public void setThroughPutSheet(XSSFSheet throughPutSheet) {
        ThroughPutSheet = throughPutSheet;
    }
}
