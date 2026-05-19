package teucost.modelos;

/**
 * Contenedor inmutable de todos los datos calculados para una visita.
 * Se construye acumulando resultados de distintas fuentes (MoveHistory,
 * Cosmos, OM Mensual) y luego se entrega al escritor para volcarlo al Excel.
 *
 * Usa un Builder para que cada fuente agregue su bloque sin depender
 * del orden de construcción.
 */
public class ResultadoNave {

    // ── Identificación ────────────────────────────────────────
    public final String nroVisita;
    public final String nombreNave;
    public final String mes;
    public final String semana;
    public final String lineaServicio;
    public final String muelle;
    public final String fechaInicio;
    public final String fechaFin;

    // ── Operación ─────────────────────────────────────────────
    public final double cuadrillas;
    public final double tiempoEfectivo;
    public final double movContenedores;
    public final double teus;

    // ── Duraciones por grúa (horas) ───────────────────────────
    public final double durSTS01, durSTS02, durSTS03;
    public final double durGM01,  durGM02;
    public final double durRTG01, durRTG02, durRTG03, durRTG04;
    public final double durRTG05, durRTG06;

    // ── Movimientos por equipo ────────────────────────────────
    public final double movRTG, movRSK, movEH;
    public final double itvAsignados;

    // ── Costos: Estiba ────────────────────────────────────────
    public final double costoUnitCuadrilla;
    public final double costoEstiba;

    // ── Costos: Diesel ────────────────────────────────────────
    public final double costoRTG;
    public final double costoRSK;
    public final double costoEH;
    public final double costoGM;
    public final double costoITV;
    public final double subtotalDiesel;

    // ── Costos: Energía eléctrica ─────────────────────────────
    public final double costoSTS01, costoSTS02, costoSTS03;
    public final double costoRTG05, costoRTG06;
    public final double subtotalEnergia;

    // ── Costos: Depreciaciones ────────────────────────────────
    public final double depSTS01, depSTS02, depSTS03;
    public final double depGM01,  depGM02;
    public final double depRTG05, depRTG06;
    public final double depRTG01, depRTG02, depRTG03, depRTG04;
    public final double depS501, depS502, depS503, depS506;
    public final double depS504, depS505, depS507;
    public final double depTT;
    public final double subtotalDepreciacion;

    // ── Resumen ───────────────────────────────────────────────
    public final double costoNaveTotal;
    public final double costoPorTeu;

    // ── Constructor privado — solo accesible desde Builder ────
    private ResultadoNave(Builder b) {
        this.nroVisita          = b.nroVisita;
        this.nombreNave         = b.nombreNave;
        this.mes                = b.mes;
        this.semana             = b.semana;
        this.lineaServicio      = b.lineaServicio;
        this.muelle             = b.muelle;
        this.fechaInicio        = b.fechaInicio;
        this.fechaFin           = b.fechaFin;

        this.cuadrillas         = b.cuadrillas;
        this.tiempoEfectivo     = b.tiempoEfectivo;
        this.movContenedores    = b.movContenedores;
        this.teus               = b.teus;

        this.durSTS01 = b.durSTS01; this.durSTS02 = b.durSTS02; this.durSTS03 = b.durSTS03;
        this.durGM01  = b.durGM01;  this.durGM02  = b.durGM02;
        this.durRTG01 = b.durRTG01; this.durRTG02 = b.durRTG02;
        this.durRTG03 = b.durRTG03; this.durRTG04 = b.durRTG04;
        this.durRTG05 = b.durRTG05; this.durRTG06 = b.durRTG06;

        this.movRTG          = b.movRTG;
        this.movRSK          = b.movRSK;
        this.movEH           = b.movEH;
        this.itvAsignados    = b.itvAsignados;

        this.costoUnitCuadrilla = b.costoUnitCuadrilla;
        this.costoEstiba        = b.costoEstiba;

        this.costoRTG        = b.costoRTG;
        this.costoRSK        = b.costoRSK;
        this.costoEH         = b.costoEH;
        this.costoGM         = b.costoGM;
        this.costoITV        = b.costoITV;
        this.subtotalDiesel  = b.subtotalDiesel;

        this.costoSTS01      = b.costoSTS01; this.costoSTS02 = b.costoSTS02;
        this.costoSTS03      = b.costoSTS03;
        this.costoRTG05      = b.costoRTG05; this.costoRTG06 = b.costoRTG06;
        this.subtotalEnergia = b.subtotalEnergia;

        this.depSTS01 = b.depSTS01; this.depSTS02 = b.depSTS02; this.depSTS03 = b.depSTS03;
        this.depGM01  = b.depGM01;  this.depGM02  = b.depGM02;
        this.depRTG05 = b.depRTG05; this.depRTG06 = b.depRTG06;
        this.depRTG01 = b.depRTG01; this.depRTG02 = b.depRTG02;
        this.depRTG03 = b.depRTG03; this.depRTG04 = b.depRTG04;
        this.depS501  = b.depS501;  this.depS502  = b.depS502;
        this.depS503  = b.depS503;  this.depS506  = b.depS506;
        this.depS504  = b.depS504;  this.depS505  = b.depS505;
        this.depS507  = b.depS507;
        this.depTT               = b.depTT;
        this.subtotalDepreciacion = b.subtotalDepreciacion;

        this.costoNaveTotal = b.costoNaveTotal;
        this.costoPorTeu    = b.costoPorTeu;
    }

    // =========================================================
    //  BUILDER
    // =========================================================

    public static class Builder {

        // Identificación
        private String nroVisita = "", nombreNave = "", mes = "", semana = "";
        private String lineaServicio = "", muelle = "", fechaInicio = "", fechaFin = "";

        // Operación
        private double cuadrillas, tiempoEfectivo, movContenedores, teus;

        // Duraciones
        private double durSTS01, durSTS02, durSTS03;
        private double durGM01, durGM02;
        private double durRTG01, durRTG02, durRTG03, durRTG04, durRTG05, durRTG06;

        // Movimientos
        private double movRTG, movRSK, movEH, itvAsignados;

        // Estiba
        private double costoUnitCuadrilla, costoEstiba;

        // Diesel
        private double costoRTG, costoRSK, costoEH, costoGM, costoITV, subtotalDiesel;

        // Energía
        private double costoSTS01, costoSTS02, costoSTS03;
        private double costoRTG05, costoRTG06, subtotalEnergia;

        // Depreciaciones
        private double depSTS01, depSTS02, depSTS03;
        private double depGM01, depGM02;
        private double depRTG05, depRTG06;
        private double depRTG01, depRTG02, depRTG03, depRTG04;
        private double depS501, depS502, depS503, depS506;
        private double depS504, depS505, depS507;
        private double depTT, subtotalDepreciacion;

        // Resumen
        private double costoNaveTotal, costoPorTeu;

        public Builder nroVisita(String v)       { this.nroVisita = v;       return this; }
        public Builder nombreNave(String v)      { this.nombreNave = v;      return this; }
        public Builder mes(String v)             { this.mes = v;             return this; }
        public Builder semana(String v)          { this.semana = v;          return this; }
        public Builder lineaServicio(String v)   { this.lineaServicio = v;   return this; }
        public Builder muelle(String v)          { this.muelle = v;          return this; }
        public Builder fechaInicio(String v)     { this.fechaInicio = v;     return this; }
        public Builder fechaFin(String v)        { this.fechaFin = v;        return this; }

        public Builder cuadrillas(double v)      { this.cuadrillas = v;      return this; }
        public Builder tiempoEfectivo(double v)  { this.tiempoEfectivo = v;  return this; }
        public Builder movContenedores(double v) { this.movContenedores = v; return this; }
        public Builder teus(double v)            { this.teus = v;            return this; }

        public Builder durSTS01(double v) { this.durSTS01 = v; return this; }
        public Builder durSTS02(double v) { this.durSTS02 = v; return this; }
        public Builder durSTS03(double v) { this.durSTS03 = v; return this; }
        public Builder durGM01(double v)  { this.durGM01  = v; return this; }
        public Builder durGM02(double v)  { this.durGM02  = v; return this; }
        public Builder durRTG01(double v) { this.durRTG01 = v; return this; }
        public Builder durRTG02(double v) { this.durRTG02 = v; return this; }
        public Builder durRTG03(double v) { this.durRTG03 = v; return this; }
        public Builder durRTG04(double v) { this.durRTG04 = v; return this; }
        public Builder durRTG05(double v) { this.durRTG05 = v; return this; }
        public Builder durRTG06(double v) { this.durRTG06 = v; return this; }

        public Builder movRTG(double v)       { this.movRTG       = v; return this; }
        public Builder movRSK(double v)       { this.movRSK       = v; return this; }
        public Builder movEH(double v)        { this.movEH        = v; return this; }
        public Builder itvAsignados(double v) { this.itvAsignados = v; return this; }

        public Builder costoUnitCuadrilla(double v) { this.costoUnitCuadrilla = v; return this; }
        public Builder costoEstiba(double v)         { this.costoEstiba        = v; return this; }

        public Builder costoRTG(double v)       { this.costoRTG       = v; return this; }
        public Builder costoRSK(double v)       { this.costoRSK       = v; return this; }
        public Builder costoEH(double v)        { this.costoEH        = v; return this; }
        public Builder costoGM(double v)        { this.costoGM        = v; return this; }
        public Builder costoITV(double v)       { this.costoITV       = v; return this; }
        public Builder subtotalDiesel(double v) { this.subtotalDiesel = v; return this; }

        public Builder costoSTS01(double v)      { this.costoSTS01      = v; return this; }
        public Builder costoSTS02(double v)      { this.costoSTS02      = v; return this; }
        public Builder costoSTS03(double v)      { this.costoSTS03      = v; return this; }
        public Builder costoRTG05(double v)      { this.costoRTG05      = v; return this; }
        public Builder costoRTG06(double v)      { this.costoRTG06      = v; return this; }
        public Builder subtotalEnergia(double v) { this.subtotalEnergia = v; return this; }

        public Builder depSTS01(double v) { this.depSTS01 = v; return this; }
        public Builder depSTS02(double v) { this.depSTS02 = v; return this; }
        public Builder depSTS03(double v) { this.depSTS03 = v; return this; }
        public Builder depGM01(double v)  { this.depGM01  = v; return this; }
        public Builder depGM02(double v)  { this.depGM02  = v; return this; }
        public Builder depRTG05(double v) { this.depRTG05 = v; return this; }
        public Builder depRTG06(double v) { this.depRTG06 = v; return this; }
        public Builder depRTG01(double v) { this.depRTG01 = v; return this; }
        public Builder depRTG02(double v) { this.depRTG02 = v; return this; }
        public Builder depRTG03(double v) { this.depRTG03 = v; return this; }
        public Builder depRTG04(double v) { this.depRTG04 = v; return this; }
        public Builder depS501(double v)  { this.depS501  = v; return this; }
        public Builder depS502(double v)  { this.depS502  = v; return this; }
        public Builder depS503(double v)  { this.depS503  = v; return this; }
        public Builder depS506(double v)  { this.depS506  = v; return this; }
        public Builder depS504(double v)  { this.depS504  = v; return this; }
        public Builder depS505(double v)  { this.depS505  = v; return this; }
        public Builder depS507(double v)  { this.depS507  = v; return this; }
        public Builder depTT(double v)               { this.depTT               = v; return this; }
        public Builder subtotalDepreciacion(double v) { this.subtotalDepreciacion = v; return this; }

        public Builder costoNaveTotal(double v) { this.costoNaveTotal = v; return this; }
        public Builder costoPorTeu(double v)    { this.costoPorTeu    = v; return this; }

        public ResultadoNave build() {
            return new ResultadoNave(this);
        }
    }

    @Override
    public String toString() {
        return String.format(
                "[Visita=%s | Nave=%s | TEUs=%.0f | Costo/TEU=%.2f]",
                nroVisita, nombreNave, teus, costoPorTeu);
    }
}
