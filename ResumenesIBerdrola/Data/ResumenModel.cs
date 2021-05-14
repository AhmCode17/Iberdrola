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
        public string Rpu { get; set; }
    }

    public class Result
    {
        public bool Success { get; set; }

        public string Error { get; set; }

        public string Msg { get; set; }
    }
}
