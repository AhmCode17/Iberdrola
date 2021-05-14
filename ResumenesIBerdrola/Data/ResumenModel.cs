using System;

namespace ResumenesIBerdrola.Data
{
    public class ResumenModel
    {
        public decimal KwhBase { get; set; }
        public decimal KwhIntermedia { get; set; }
        public decimal KwhPunta { get; set; }
        public decimal KwhSemiPunta { get; set; }
        public decimal KwhTotales { get; set; }
        public decimal KwBase { get; set; }
        public decimal KwIntermedia { get; set; }
        public decimal KwPunta { get; set; }
        public decimal KwSemiPunta { get; set; }
        public decimal KwKvarh { get; set; }
        public decimal KwFp { get; set; }
        public string Tipo { get; set; }
        public string Concepto { get; set; }
        public string NombreCliente { get; set; }
        public string Periodo { get; set; }
        public string Rpu { get; set; }

        public int FkResumen { get; set; }

        public int FkConcepto { get; set; }
        public decimal CapacidadTotal { get; set; }
        public decimal DistribucionTotal { get; set; }
        public string Descripcion { get; set; }

    }

    public class Result
    {
        public bool Success { get; set; }

        public string Error { get; set; }

        public string Msg { get; set; }
        public object Data { get; set; }
    }

    public class ConceptoModel
    {
        public int Id { get; set; }
        public string Concepto { get; set; }
        public string TipoConcepto { get; set; }
    }

    public class CentralModel
    {
        public int Id { get; set; }
        public string RazonSocial { get; set; }
        public string Cliente { get; set; }
        public string Planta { get; set; }
        public string Nombre { get; set; }
        public string Cfe { get; set; }
    }

    public class ResumenBaseModel
    {
        public int Id { get; set; }
        public int FkCentral { get; set; }
        public string Periodo { get; set; }
        public DateTime FechaCreacion { get; set; }
    }
}
