package cuadrillas;

public class ControladorCuadrillas {

    private String rutaCarpetaNaves;
    private String rutaExcelDestino;

    public ControladorCuadrillas(
            String rutaCarpetaNaves,
            String rutaExcelDestino
    ) {
        this.rutaCarpetaNaves = rutaCarpetaNaves;
        this.rutaExcelDestino = rutaExcelDestino;
    }

    public void extraerDatosCuadrillas() {

        // Carpeta temporal donde se copiarán los PDFs SAP antes de procesarlos
        String rutaTemporal = "C:\\Temp\\CuadrillasTPE\\";

        // Paso 1: Copiar los PDFs SAP desde la carpeta de naves a la carpeta temporal
        LecturaArchivosSAP lecturaArchivosSAP = new LecturaArchivosSAP();
        lecturaArchivosSAP.setRutaCarpetaNaves(this.rutaCarpetaNaves);
        lecturaArchivosSAP.copiarPdfsADestino(rutaTemporal);

        // Paso 2: Aplicar OCR sobre los PDFs copiados y extraer datos al Excel destino
        ExtraccionCuadrillas extraccionCuadrillas = new ExtraccionCuadrillas();
        extraccionCuadrillas.extraerDatosAExcel(
                rutaTemporal,
                this.rutaExcelDestino
        );
    }
}