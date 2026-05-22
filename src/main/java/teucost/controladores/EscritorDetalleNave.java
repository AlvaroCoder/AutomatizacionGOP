package teucost.controladores;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import teucost.modelos.LecturaExcels;

import java.util.HashMap;


public class EscritorDetalleNave {

    private final String rutaExcelDestino;
    private final int    hojaDetalle = 1;


    private double tasaDepreciacionGM = 45.54;

    private static final int D_MES       = 0;
    private static final int D_SEMANA    = 1;
    private static final int D_NAVE      = 2;
    private static final int D_VISITA    = 3;
    private static final int D_LINEA     = 4;
    private static final int D_MUELLE    = 5;
    private static final int D_FECHA_FIN = 6;
    private static final int D_CUADRILLAS = 7;
    private static final int D_T_EFECTIVO = 8;
    private static final int D_MOV_CONT   = 9;
    private static final int D_TEUS       = 10;

    private static final int D_DUR_STS01 = 11;
    private static final int D_DUR_STS02 = 12;
    private static final int D_DUR_STS03 = 13;
    private static final int D_DUR_GM01  = 14;
    private static final int D_DUR_GM02  = 15;
    private static final int D_DUR_RTG01 = 16;
    private static final int D_DUR_RTG02 = 17;
    private static final int D_DUR_RTG03 = 18;
    private static final int D_DUR_RTG04 = 19;
    private static final int D_DUR_RTG05 = 20;
    private static final int D_DUR_RTG06 = 21;
    private static final int D_DUR_S501 = 22;
    private static final int D_DUR_S502 = 23;
    private static final int D_DUR_S503 = 24;
    private static final int D_DUR_S504 = 25;
    private static final int D_DUR_S505 = 26;
    private static final int D_DUR_S506 = 27;
    private static final int D_DUR_S507 = 28;

    private static final int D_MOV_STS01  = 29;
    private static final int D_MOV_STS02  = 30;
    private static final int D_MOV_STS03  = 31;
    private static final int D_MOV_RTG01  = 32;
    private static final int D_MOV_RTG02  = 33;
    private static final int D_MOV_RTG03  = 34;
    private static final int D_MOV_RTG04  = 35;
    private static final int D_MOV_RTG05  = 36;
    private static final int D_MOV_RTG06  = 37;
    private static final int D_MOV_S501 = 38;
    private  static final int D_MOV_S502 = 39;
    private static final int D_MOV_S503 = 40;
    private static final int D_MOV_S504 = 41;
    private static final int D_MOV_S505 = 42;
    private static final int D_MOV_S506 = 43;
    private static final int D_MOV_S507 = 44;

    private static final int D_EST_UNIT  = 45;
    private static final int D_EST_TOTAL = 46;

    private static final int D_RTG_COSTO = 47;
    private static final int D_RSK_COSTO = 48;
    private static final int D_EH_COSTO  = 49;
    private static final int D_GM_COSTO  = 50;
    private static final int D_ITV_COSTO = 51;
    private static final int D_DIESEL_SUB= 52;

    private static final int D_STS01_COSTO = 53;
    private static final int D_STS02_COSTO = 54;
    private static final int D_STS03_COSTO = 55;
    private static final int D_RTG05_COSTO = 56;
    private static final int D_RTG06_COSTO = 57;
    private static final int D_ENER_SUB    = 58;

    private static final int D_DEP_STS01 = 59;
    private static final int D_DEP_STS02 = 60;
    private static final int D_DEP_STS03 = 61;
    private static final int D_DEP_GM01  = 62;
    private static final int D_DEP_GM02  = 63;
    private static final int D_DEP_RTG05 = 64;
    private static final int D_DEP_RTG06 = 65;
    private static final int D_DEP_RTG01 = 66;
    private static final int D_DEP_RTG02 = 67;
    private static final int D_DEP_RTG03 = 68;
    private static final int D_DEP_RTG04 = 69;
    private static final int D_DEP_S501  = 70;
    private static final int D_DEP_S502  = 71;
    private static final int D_DEP_S503  = 72;
    private static final int D_DEP_S506  = 73;
    private static final int D_DEP_S504  = 74;
    private static final int D_DEP_S505  = 75;
    private static final int D_DEP_S507  = 76;
    private static final int D_DEP_TT    = 77;
    private static final int D_DEP_SUB   = 78;

    private static final int D_RES_ESTIBA  = 79;
    private static final int D_RES_DIESEL  = 80;
    private static final int D_RES_ENERGIA = 81;
    private static final int D_RES_DEPREC  = 82;
    private static final int D_RES_TOTAL   = 83;
    private static final int D_RES_TEUS    = 84;
    private static final int D_COSTO_TEU   = 85;

    private static final Object[][] GRUPOS = {
            { "IDENTIFICACIÓN",    D_MES,         D_FECHA_FIN,   "C9B1FF" },
            { "OPERACIÓN",         D_CUADRILLAS,  D_TEUS,        "9FE1CB" },
            { "DURACIONES (hrs)",  D_DUR_STS01,   D_DUR_S507,   "B5D4F4" },
            {"MOVIMIENTOS (FETCH y PUT)", D_MOV_STS01, D_MOV_S507, "D3D1C7" },
            { "ESTIBA",            D_EST_UNIT,    D_EST_TOTAL,   "FAC775" },
            { "DIESEL",            D_RTG_COSTO,     D_DIESEL_SUB,  "F5C4B3" },
            { "ENERGÍA ELÉCTRICA", D_STS01_COSTO, D_ENER_SUB,    "C0DD97" },
            { "DEPRECIACIONES",    D_DEP_STS01,   D_DEP_SUB,     "D3D1C7" },
            { "RESUMEN",           D_RES_ESTIBA,  D_COSTO_TEU,   "F4C0D1" },
    };

    private static final String[] NOMBRES_COLUMNAS = {
            // Identificación
            "Mes", "Semana", "Nombre nave", "Nro visita",
            "Línea servicio", "Muelle", "Fecha fin",
            // Operación
            "Cuadrillas", "T. efectivo (hrs)", "Mov. contenedores", "TEUs",
            // Duraciones
            "STS01 (hrs)", "STS02 (hrs)", "STS03 (hrs)",
            "GM01 (hrs)",  "GM02 (hrs)",
            "RTG01 (hrs)", "RTG02 (hrs)", "RTG03 (hrs)", "RTG04 (hrs)",
            "RTG05 (hrs)", "RTG06 (hrs)",
            "S-501 (hrs)", "S-502 (hrs)", "S-503 (hrs)", "S-504 (hrs)", "S-505 (hrs)",
            "S-506 (hrs)", "S-507 (hrs)",
            // Movimientos
            "STS01 (mov)", "STS02 (mov)", "STS03 (mov)",
            "RTG01 (mov)", "RTG02 (mov)", "RTG03 (mov)", "RTG04 (mov)","RTG05 (mov)",
            "RTG06 (mov)",
            "S-501 (mov)", "S-502 (mov)","S-503 (mov)", "S-504 (mov)",
            "S-505 (mov)", "S-506 (mov)","S-507 (mov)",
            // Estiba
            "C. unit. cuadrilla", "Total estiba",
            // Diesel
            "RTG - costo ($)",
            "RSK - costo ($)",
            "EH - costo ($)",
            "GM - costo ($)",
            "ITV - costo ($)",
            "Subtotal diesel",
            // Energía
            "STS01 - costo ($)", "STS02 - costo ($)", "STS03 - costo ($)",
            "RTG05 - costo ($)", "RTG06 - costo ($)",
            "Subtotal energía",
            // Depreciaciones
            "Dep. STS01", "Dep. STS02", "Dep. STS03",
            "Dep. GM01",  "Dep. GM02",
            "Dep. RTG05", "Dep. RTG06",
            "Dep. RTG01", "Dep. RTG02", "Dep. RTG03", "Dep. RTG04",
            "Dep. S-501", "Dep. S-502", "Dep. S-503", "Dep. S-506",
            "Dep. S-504", "Dep. S-505", "Dep. S-507",
            "Dep. TT (ITV)",
            "Subtotal deprec.",
            // Resumen
            "Estiba total", "Diesel total", "Energía total",
            "Depreciación total", "COSTO NAVE TOTAL", "TEUs", "COSTO / TEU"
    };


    public EscritorDetalleNave(String rutaExcelDestino) {
        this.rutaExcelDestino = rutaExcelDestino;
    }

    public void escribirFila(HashMap<String, Object> datos) {
        if (datos == null || datos.isEmpty()) {
            System.out.println("[WARN] EscritorDetalleNave: HashMap vacío, se omite la escritura.");
            return;
        }

        LecturaExcels excel = null;
        try {
            excel = new LecturaExcels(this.rutaExcelDestino);

            // Garantizar que la hoja 1 exista antes de operar sobre ella
            excel.garantizarHoja(hojaDetalle, "Detalle Costos");

            // Crear encabezados solo la primera vez
            if (!excel.hojaContieneEncabezados(hojaDetalle, 0)) {
                crearEncabezados(excel);
            }

            // ← CORRECCIÓN: la hoja de detalle es plana, sin XSSFTable formal.
            //   Usar agregarFilaSinTabla en lugar de insertarFilaEnTabla.
            Row fila = excel.agregarFilaSinTabla(hojaDetalle, 0);

            poblarFila(fila, datos);

            excel.guardar();
            System.out.println("[OK] EscritorDetalleNave → fila escrita para visita: "
                    + str(datos, "nroVisita"));

        } catch (Exception e) {
            System.out.println("[ERROR] EscritorDetalleNave.escribirFila: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (excel != null) try { excel.cerrar(); } catch (Exception ignored) {}
        }
    }

    private void poblarFila(Row fila, HashMap<String, Object> d) {

        txt(fila, D_MES,       str(d, "mes"));
        txt(fila, D_SEMANA,    str(d, "semana"));
        txt(fila, D_NAVE,      str(d, "nombreVisita"));
        txt(fila, D_VISITA,    str(d, "nroVisita"));
        txt(fila, D_LINEA,     str(d, "lineaServicio"));
        txt(fila, D_MUELLE,    str(d, "muelle"));
        txt(fila, D_FECHA_FIN, str(d, "fechaFin"));

        num(fila, D_CUADRILLAS, dbl(d, "cuadrillas"));
        num(fila, D_T_EFECTIVO, dbl(d, "tiempoEfectivo"));
        num(fila, D_MOV_CONT,   dbl(d, "movCont"));
        num(fila, D_TEUS,       dbl(d, "teus"));

        num(fila, D_DUR_STS01, dbl(d, "duracionSTS01"));
        num(fila, D_DUR_STS02, dbl(d, "duracionSTS02"));
        num(fila, D_DUR_STS03, dbl(d, "duracionSTS03"));

        double durGM01 = tasaDepreciacionGM > 0
                ? dbl(d, "costoDepreciacionGM01") / tasaDepreciacionGM : 0;
        double durGM02 = tasaDepreciacionGM > 0
                ? dbl(d, "costoDepreciacionGM02") / tasaDepreciacionGM : 0;
        num(fila, D_DUR_GM01,  durGM01);
        num(fila, D_DUR_GM02,  durGM02);
        num(fila, D_DUR_RTG01, dbl(d, "duracionRTG01"));
        num(fila, D_DUR_RTG02, dbl(d, "duracionRTG02"));
        num(fila, D_DUR_RTG03, dbl(d, "duracionRTG03"));
        num(fila, D_DUR_RTG04, dbl(d, "duracionRTG04"));
        num(fila, D_DUR_RTG05, dbl(d, "duracionRTG05"));
        num(fila, D_DUR_RTG06, dbl(d, "duracionRTG06"));

        num(fila, D_DUR_S501, dbl(d, "duracionS501"));
        num(fila, D_DUR_S502, dbl(d, "duracionS502"));
        num(fila, D_DUR_S503, dbl(d, "duracionS503"));
        num(fila, D_DUR_S504, dbl(d, "duracionS504"));
        num(fila, D_DUR_S505, dbl(d, "duracionS505"));
        num(fila, D_DUR_S506, dbl(d, "duracionS506"));
        num(fila, D_DUR_S507, dbl(d, "duracionS507"));

        num(fila, D_MOV_STS01, dbl(d, "movimientosSTS01"));
        num(fila, D_MOV_STS02, dbl(d, "movimientosSTS02"));
        num(fila, D_MOV_STS03, dbl(d, "movimientosSTS03"));

        num(fila, D_MOV_RTG01, dbl(d, "movimientosRTG01"));
        num(fila, D_MOV_RTG02, dbl(d, "movimientosRTG02"));
        num(fila, D_MOV_RTG03, dbl(d, "movimientosRTG03"));
        num(fila, D_MOV_RTG04, dbl(d, "movimientosRTG04"));
        num(fila, D_MOV_RTG05, dbl(d, "movimientosRTG05"));
        num(fila, D_MOV_RTG06, dbl(d, "movimientosRTG06"));

        num(fila, D_MOV_S501, dbl(d, "movimientosS501"));
        num(fila, D_MOV_S502, dbl(d, "movimientosS502"));
        num(fila, D_MOV_S503, dbl(d, "movimientosS503"));
        num(fila, D_MOV_S504, dbl(d, "movimientosS504"));
        num(fila, D_MOV_S505, dbl(d, "movimientosS505"));
        num(fila, D_MOV_S506, dbl(d, "movimientosS506"));
        num(fila, D_MOV_S507, dbl(d, "movimientosS507"));

        double cuadrillas  = dbl(d, "cuadrillas");
        double totalEstiba = dbl(d, "costoEstiba");
        double unitEstiba  = cuadrillas > 0 ? totalEstiba / cuadrillas : 0;
        num(fila, D_EST_UNIT,  unitEstiba);
        num(fila, D_EST_TOTAL, totalEstiba);

        num(fila, D_RTG_COSTO,  dbl(d, "costoRTG"));
        num(fila, D_RSK_COSTO,  dbl(d, "costoReactStacker"));
        num(fila, D_EH_COSTO,   dbl(d, "costoEmptyHand"));
        num(fila, D_GM_COSTO,   dbl(d, "costoTotalGruaMovil"));
        num(fila, D_ITV_COSTO,  dbl(d, "costoTotalITV"));
        num(fila, D_DIESEL_SUB, dbl(d, "subtotalCosto"));

        // ── ENERGÍA ELÉCTRICA ─────────────────────────────────────
        num(fila, D_STS01_COSTO, dbl(d, "costoTotalSTS01"));
        num(fila, D_STS02_COSTO, dbl(d, "costoTotalSTS02"));
        num(fila, D_STS03_COSTO, dbl(d, "costoTotalSTS03"));
        num(fila, D_RTG05_COSTO, dbl(d, "costoTotalRTG05"));
        num(fila, D_RTG06_COSTO, dbl(d, "costoTotalRTG06"));
        num(fila, D_ENER_SUB,    dbl(d, "costoSubtotalSTSRTG"));

        // ── DEPRECIACIONES ────────────────────────────────────────
        num(fila, D_DEP_STS01, dbl(d, "costoDepreciacionSTS01"));
        num(fila, D_DEP_STS02, dbl(d, "costoDepreciacionSTS02"));
        num(fila, D_DEP_STS03, dbl(d, "costoDepreciacionSTS03"));
        num(fila, D_DEP_GM01,  dbl(d, "costoDepreciacionGM01"));
        num(fila, D_DEP_GM02,  dbl(d, "costoDepreciacionGM02"));
        num(fila, D_DEP_RTG05, dbl(d, "costoDepreciacionRTG05"));
        num(fila, D_DEP_RTG06, dbl(d, "costoDepreciacionRTG06"));
        num(fila, D_DEP_RTG01, dbl(d, "costoDepreciacionRTG01"));
        num(fila, D_DEP_RTG02, dbl(d, "costoDepreciacionRTG02"));
        num(fila, D_DEP_RTG03, dbl(d, "costoDepreciacionRTG03"));
        num(fila, D_DEP_RTG04, dbl(d, "costoDepreciacionRTG04"));
        num(fila, D_DEP_S501,  dbl(d, "costoDepreciacionS501"));
        num(fila, D_DEP_S502,  dbl(d, "costoDepreciacionS502"));
        num(fila, D_DEP_S503,  dbl(d, "costoDepreciacionS503"));
        num(fila, D_DEP_S506,  dbl(d, "costoDepreciacionS506"));
        num(fila, D_DEP_S504,  dbl(d, "costoDepreciacionS504"));
        num(fila, D_DEP_S505,  dbl(d, "costoDepreciacionS505"));
        num(fila, D_DEP_S507,  dbl(d, "costoDepreciacionS507"));
        num(fila, D_DEP_TT,    dbl(d, "costoDepreciacionTT"));
        num(fila, D_DEP_SUB,   dbl(d, "subtotalCostoDepreciacion"));

        // ── RESUMEN FINAL ─────────────────────────────────────────
        num(fila, D_RES_ESTIBA,  totalEstiba);
        num(fila, D_RES_DIESEL,  dbl(d, "subtotalCosto"));
        num(fila, D_RES_ENERGIA, dbl(d, "costoSubtotalSTSRTG"));
        num(fila, D_RES_DEPREC,  dbl(d, "subtotalCostoDepreciacion"));
        num(fila, D_RES_TOTAL,   dbl(d, "costoNaveOperativo"));
        num(fila, D_RES_TEUS,    dbl(d, "teus"));
        num(fila, D_COSTO_TEU,   dbl(d, "costoPorTeu"));
    }

    private boolean hojaContieneEncabezados(LecturaExcels excel) {
        try {
            Sheet hoja = excel.getWorkbook().getSheetAt(hojaDetalle);
            if (hoja == null || hoja.getPhysicalNumberOfRows() == 0) return false;
            Row primera = hoja.getRow(0);
            if (primera == null) return false;
            Cell celda = primera.getCell(0);
            return celda != null && celda.getCellType() != CellType.BLANK;
        } catch (Exception e) {
            return false;
        }
    }

    private void crearEncabezados(LecturaExcels excel) {
        try {
            Workbook wb   = excel.getWorkbook();
            Sheet    hoja = wb.getSheetAt(hojaDetalle);
            if (hoja == null) hoja = wb.createSheet("Detalle Costos");

            // ── Fila 0: grupos con color de fondo ─────────────────
            Row filaGrupos = hoja.createRow(0);
            filaGrupos.setHeightInPoints(22);

            for (Object[] g : GRUPOS) {
                String nombre   = (String) g[0];
                int    colIni   = (int)    g[1];
                int    colFin   = (int)    g[2];
                String hexColor = (String) g[3];

                CellStyle estilo = wb.createCellStyle();
                estilo.setFillForegroundColor(new XSSFColor(
                        hexToBytes(hexColor), new DefaultIndexedColorMap()));
                estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                estilo.setAlignment(HorizontalAlignment.CENTER);
                estilo.setVerticalAlignment(VerticalAlignment.CENTER);
                Font font = wb.createFont();
                font.setBold(true);
                font.setFontHeightInPoints((short) 10);
                estilo.setFont(font);

                // Celda de inicio + merge
                Cell celda = filaGrupos.createCell(colIni);
                celda.setCellValue(nombre);
                celda.setCellStyle(estilo);

                // Rellenar el resto del grupo con el mismo estilo
                for (int c = colIni + 1; c <= colFin; c++) {
                    filaGrupos.createCell(c).setCellStyle(estilo);
                }
                if (colFin > colIni) {
                    hoja.addMergedRegion(
                            new CellRangeAddress(0, 0, colIni, colFin));
                }
            }

            Row filaNombres = hoja.createRow(1);
            filaNombres.setHeightInPoints(32);

            CellStyle estiloCol = wb.createCellStyle();
            estiloCol.setFillForegroundColor(new XSSFColor(
                    new byte[]{(byte)0xF1,(byte)0xEF,(byte)0xE8},
                    new DefaultIndexedColorMap()));
            estiloCol.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estiloCol.setAlignment(HorizontalAlignment.CENTER);
            estiloCol.setVerticalAlignment(VerticalAlignment.CENTER);
            estiloCol.setWrapText(true);
            Font fontCol = wb.createFont();
            fontCol.setFontHeightInPoints((short) 9);
            estiloCol.setFont(fontCol);

            for (int i = 0; i < NOMBRES_COLUMNAS.length; i++) {
                Cell c = filaNombres.createCell(i);
                c.setCellValue(NOMBRES_COLUMNAS[i]);
                c.setCellStyle(estiloCol);
                hoja.setColumnWidth(i, 4400);   // ≈ 14–15 chars
            }

            System.out.println("[OK] EscritorDetalleNave → encabezados creados en hoja 1.");

        } catch (Exception e) {
            System.out.println("[WARN] EscritorDetalleNave → no se pudieron crear encabezados: "
                    + e.getMessage());
        }
    }

    private void txt(Row fila, int col, String valor) {
        Cell c = fila.getCell(col);
        if (c == null) c = fila.createCell(col);
        c.setCellValue(valor != null ? valor : "");
    }

    private void num(Row fila, int col, double valor) {
        Cell c = fila.getCell(col);
        if (c == null) c = fila.createCell(col);
        c.setCellValue(valor);
    }

    private String str(HashMap<String, Object> d, String key) {
        Object v = d.get(key);
        return v != null ? v.toString() : "";
    }

    private double dbl(HashMap<String, Object> d, String key) {
        Object v = d.get(key);
        if (v instanceof Double)  return (Double) v;
        if (v instanceof Integer) return ((Integer) v).doubleValue();
        return 0.0;   // clave ausente → 0, sin lanzar excepción
    }

    private byte[] hexToBytes(String hex) {
        return new byte[]{
                (byte) Integer.parseInt(hex.substring(0, 2), 16),
                (byte) Integer.parseInt(hex.substring(2, 4), 16),
                (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }

    public void setTasaDepreciacionGM(double tasa) {
        this.tasaDepreciacionGM = tasa;
    }
}