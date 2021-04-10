using System;
using System.Collections.Generic;

namespace EntregasRendir.Context.Models
{
    public partial class EntregaRendir
    {
        public string CodigoEmpresa { get; set; }
        public string ComprobanteTesoreria { get; set; }
        public string Moneda { get; set; }
        public string CorrelativoHelm { get; set; }
        public DateTime? Fecha { get; set; }
        public decimal? TipoCambio { get; set; }
        public string Glosa { get; set; }
        public string SubCuenta { get; set; }
        public string TipoConcepto { get; set; }
        public string CodigoConcepto { get; set; }
        public decimal? Importe { get; set; }
        public string EstatusValidacion { get; set; }
        public int Secuencial { get; set; }
        public string RegistradoPor { get; set; }
        public string ObservacionValidacion { get; set; }
    }
}
