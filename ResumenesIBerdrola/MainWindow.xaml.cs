using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using ResumenesIBerdrola.Data;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
namespace ResumenesIBerdrola
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] tipos = { "ENERGIA TOTAL", "ENERGIA NORMAL", "ENERGIA PORTEADA", "ENERGIA NORMAL POR FALTANTE", "RESPALDO POR CARGA" };
        public MainWindow()
        {
            InitializeComponent();
        }
        private IDbConnection _conn;
        public IDbConnection Connection
        {
            get
            {
                return _conn;
            }
        }

        private void BtnSalir_Click(object sender, RoutedEventArgs e)
        {
          

            //this.Close();
          
          



        }

        private void BtnSeleccionar_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                InitialDirectory = @"C:\Descargas",
                IsFolderPicker = true
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtRuta.Text = dialog.FileName;

                foreach (string file in Directory.EnumerateFiles(txtRuta.Text, "*.xlsx"))
                {
                    lstFiles.Items.Add(file);
                }

                foreach (string file in Directory.EnumerateFiles(txtRuta.Text, "*.xls"))
                {
                    lstFiles.Items.Add(file);
                    lblFiles.Content = lstFiles.Items.Count + " archivos";
                }
            }
        }

        private void BtnProcesar_Click(object sender, RoutedEventArgs e)
        {
            //GetHeaderExcel();
            GetHeaderExcelOld();
        }

        public void GetHeaderExcel()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileInfo existingFile = new FileInfo(@"C:\Users\Babel\Documents\Iberdrola\Resumen CFE\Resumen BNS.xlsx");
            List<ResumenModel> lst = new List<ResumenModel>();
            List<ResumenModel> lstDetail = new List<ResumenModel>();
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                var currentSheet = package.Workbook.Worksheets;

                ExcelWorksheet worksheet = currentSheet.First();
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                var perido = worksheet.Cells[4, 1].Value.ToString().Split(':')[0];
                var central = worksheet.Cells[3, 1].Value.ToString().Split(':')[1];
                var tipo = worksheet.Cells[17, 2].Value.ToString();
                int endLetter = perido.Length - 7;
                perido = perido.Substring(0, endLetter);



                int row2 = 17;

                ///Obtiene los headers
                for (int row = 7; row <= 16; row++)
                {
                    var concepto = worksheet.Cells[row, 2].Value.ToString();
                    var KwhBase = worksheet.Cells[row, 3].Value == null ? string.Empty : worksheet.Cells[row, 3].Value.ToString();
                    var KwhIntermedia = worksheet.Cells[row, 4].Value == null ? string.Empty : worksheet.Cells[row, 4].Value.ToString();
                    var KwhPunta = worksheet.Cells[row, 5].Value == null ? string.Empty : worksheet.Cells[row, 5].Value.ToString();
                    var KwhSemiPunta = worksheet.Cells[row, 6].Value == null ? string.Empty : worksheet.Cells[row, 6].Value.ToString();
                    var KwhTotales = worksheet.Cells[row, 7].Value == null ? string.Empty : worksheet.Cells[row, 7].Value.ToString();
                    var KwBase = worksheet.Cells[row, 8].Value == null ? string.Empty : worksheet.Cells[row, 8].Value.ToString();
                    var KwIntermedia = worksheet.Cells[row, 9].Value == null ? string.Empty : worksheet.Cells[row, 9].Value.ToString();
                    var KwPunta = worksheet.Cells[row, 10].Value == null ? string.Empty : worksheet.Cells[row, 10].Value.ToString();
                    var KwSemiPunta = worksheet.Cells[row, 11].Value == null ? string.Empty : worksheet.Cells[row, 11].Value.ToString();
                    var KwKvarh = worksheet.Cells[row, 12].Value == null ? string.Empty : worksheet.Cells[row, 12].Value.ToString();
                    var KwFp = worksheet.Cells[row, 13].Value == null ? string.Empty : worksheet.Cells[row, 13].Value.ToString();

                    lst.Add(new ResumenModel
                    {
                        KwhBase = decimal.Parse(KwhBase),
                        KwhIntermedia = decimal.Parse(KwhIntermedia),
                        KwhPunta = decimal.Parse(KwhPunta),
                        KwhSemiPunta = decimal.Parse(KwhSemiPunta),
                        KwhTotales = decimal.Parse(KwhTotales),
                        KwBase = decimal.Parse(KwBase),
                        KwIntermedia = decimal.Parse(KwIntermedia),
                        KwPunta = decimal.Parse(KwPunta),
                        KwSemiPunta = decimal.Parse(KwSemiPunta),
                        KwKvarh = decimal.Parse(KwKvarh),
                        KwFp = decimal.Parse(KwFp),
                        Concepto = concepto

                    });
                }
                //Obtiene el detalle
                lstDetail.AddRange(GetDetail(worksheet, row2, rowCount, tipos, tipo));
            }
        }

        public List<ResumenModel> GetDetail(ExcelWorksheet worksheet, int row2, int rowCount, string[] tipos, string tipo)
        {
            List<ResumenModel> lst = new List<ResumenModel>();
            for (int row = row2; row <= rowCount; row++)
            {
                if (tipos.Contains(worksheet.Cells[row, 2].Value.ToString()))
                {
                    lst.AddRange(GetDetail(worksheet, row + 1, rowCount, tipos, worksheet.Cells[row, 2].Value.ToString()));
                    break;
                }
                else
                {
                    var concepto = worksheet.Cells[row, 2].Value.ToString();
                    var KwhBase = worksheet.Cells[row, 3].Value == null ? string.Empty : worksheet.Cells[row, 3].Value.ToString();
                    var KwhIntermedia = worksheet.Cells[row, 4].Value == null ? string.Empty : worksheet.Cells[row, 4].Value.ToString();
                    var KwhPunta = worksheet.Cells[row, 5].Value == null ? string.Empty : worksheet.Cells[row, 5].Value.ToString();
                    var KwhSemiPunta = worksheet.Cells[row, 6].Value == null ? string.Empty : worksheet.Cells[row, 6].Value.ToString();
                    var KwhTotales = worksheet.Cells[row, 7].Value == null ? string.Empty : worksheet.Cells[row, 7].Value.ToString();
                    var KwBase = worksheet.Cells[row, 8].Value == null ? string.Empty : worksheet.Cells[row, 8].Value.ToString();
                    var KwIntermedia = worksheet.Cells[row, 9].Value == null ? string.Empty : worksheet.Cells[row, 9].Value.ToString();
                    var KwPunta = worksheet.Cells[row, 10].Value == null ? string.Empty : worksheet.Cells[row, 10].Value.ToString();
                    var KwSemiPunta = worksheet.Cells[row, 11].Value == null ? string.Empty : worksheet.Cells[row, 11].Value.ToString();
                    var KwKvarh = worksheet.Cells[row, 12].Value == null ? string.Empty : worksheet.Cells[row, 12].Value.ToString();
                    var KwFp = worksheet.Cells[row, 13].Value == null ? string.Empty : worksheet.Cells[row, 13].Value.ToString();
                    var nomCliente = concepto.Split('_')[2];
                    var rpu = concepto.Split('_')[1];
                    lst.Add(new ResumenModel
                    {
                        KwhBase = decimal.Parse(KwhBase),
                        KwhIntermedia = decimal.Parse(KwhIntermedia),
                        KwhPunta = decimal.Parse(KwhPunta),
                        KwhSemiPunta = decimal.Parse(KwhSemiPunta),
                        KwhTotales = decimal.Parse(KwhTotales),
                        KwBase = decimal.Parse(KwBase),
                        KwIntermedia = decimal.Parse(KwIntermedia),
                        KwPunta = decimal.Parse(KwPunta),
                        KwSemiPunta = decimal.Parse(KwSemiPunta),
                        KwKvarh = decimal.Parse(KwKvarh),
                        KwFp = decimal.Parse(KwFp),
                        Tipo = tipo,
                        Concepto = concepto,
                        NombreCliente = nomCliente,
                        Rpu = rpu
                    });
                }
            }
            return lst;
        }

        public void GetHeaderExcelOld()
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            ISheet sheet;
            List<ResumenModel> lst = new List<ResumenModel>();
            List<ResumenModel> lstDetail = new List<ResumenModel>();
            using (var stream = new FileStream(@"C:\Users\Babel\Documents\Iberdrola\Resumen CFE\Resumen CDUII.xls", FileMode.Open))
            {
                stream.Position = 0;
                HSSFWorkbook xssWorkbook = new HSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;

                var central = string.Empty;
                var perido = string.Empty;
                var tipo = string.Empty;

                IRow centralRow = sheet.GetRow(2);
                ICell cellCentral = centralRow.GetCell(0);

                if (cellCentral != null || !string.IsNullOrWhiteSpace(cellCentral.ToString()))
                    central = cellCentral.ToString().Split(':')[1];

                centralRow = sheet.GetRow(3);
                cellCentral = centralRow.GetCell(0);

                if (cellCentral != null || !string.IsNullOrWhiteSpace(cellCentral.ToString()))
                    perido = cellCentral.ToString().Split(':')[0];

                centralRow = sheet.GetRow(16);
                cellCentral = centralRow.GetCell(1);
                if (cellCentral != null || !string.IsNullOrWhiteSpace(cellCentral.ToString()))
                    tipo = cellCentral.ToString();

                int endLetter = perido.Length - 7;
                perido = perido.Substring(0, endLetter);


                for (int i = 6; i <= 15; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                    var concepto = row.GetCell(1).ToString();
                    var KwhBase = row.GetCell(2).ToString();
                    var KwhIntermedia = row.GetCell(3).ToString();
                    var KwhPunta = row.GetCell(4).ToString();
                    var KwhSemiPunta = row.GetCell(5).ToString();
                    var KwhTotales = row.GetCell(6).ToString();
                    var KwBase = row.GetCell(7).ToString();
                    var KwIntermedia = row.GetCell(8).ToString();
                    var KwPunta = row.GetCell(9).ToString();
                    var KwSemiPunta = row.GetCell(10).ToString();
                    var KwKvarh = row.GetCell(11).ToString();
                    var KwFp = row.GetCell(12).ToString();

                    lst.Add(new ResumenModel
                    {
                        KwhBase = decimal.Parse(KwhBase),
                        KwhIntermedia = decimal.Parse(KwhIntermedia),
                        KwhPunta = decimal.Parse(KwhPunta),
                        KwhSemiPunta = decimal.Parse(KwhSemiPunta),
                        KwhTotales = decimal.Parse(KwhTotales),
                        KwBase = decimal.Parse(KwBase),
                        KwIntermedia = decimal.Parse(KwIntermedia),
                        KwPunta = decimal.Parse(KwPunta),
                        KwSemiPunta = decimal.Parse(KwSemiPunta),
                        KwKvarh = decimal.Parse(KwKvarh),
                        KwFp = decimal.Parse(KwFp),
                        Concepto = concepto
                    });

                    //Obtiene el detalle
                    lstDetail.AddRange(GetDetailOld(sheet, 16, sheet.LastRowNum, tipos, tipo));
                }
            }
        }

        public List<ResumenModel> GetDetailOld(ISheet sheet, int row2, int rowCount, string[] tipos, string tipo)
        {
            List<ResumenModel> lst = new List<ResumenModel>();

            for (int row = row2; row <= rowCount; row++)
            {
                IRow worksheet = sheet.GetRow(row);
                if (tipos.Contains(worksheet.GetCell(1).ToString()))
                {
                    lst.AddRange(GetDetailOld(sheet, row + 1, rowCount, tipos, worksheet.GetCell(1).ToString()));
                    break;
                }
                else
                {
                    var concepto = worksheet.GetCell(1).ToString();
                    var KwhBase = worksheet.GetCell(2).ToString();
                    var KwhIntermedia = worksheet.GetCell(3).ToString();
                    var KwhPunta = worksheet.GetCell(4).ToString();
                    var KwhSemiPunta = worksheet.GetCell(5).ToString();
                    var KwhTotales = worksheet.GetCell(6).ToString();
                    var KwBase = worksheet.GetCell(7).ToString();
                    var KwIntermedia = worksheet.GetCell(8).ToString();
                    var KwPunta = worksheet.GetCell(9).ToString();
                    var KwSemiPunta = worksheet.GetCell(10).ToString();
                    var KwKvarh = worksheet.GetCell(11).ToString();
                    var KwFp = worksheet.GetCell(12).ToString();
                    var nomCliente = concepto.Split('_')[2];
                    var rpu = concepto.Split('_')[1];
                    lst.Add(new ResumenModel
                    {
                        KwhBase = decimal.Parse(KwhBase),
                        KwhIntermedia = decimal.Parse(KwhIntermedia),
                        KwhPunta = decimal.Parse(KwhPunta),
                        KwhSemiPunta = decimal.Parse(KwhSemiPunta),
                        KwhTotales = decimal.Parse(KwhTotales),
                        KwBase = decimal.Parse(KwBase),
                        KwIntermedia = decimal.Parse(KwIntermedia),
                        KwPunta = decimal.Parse(KwPunta),
                        KwSemiPunta = decimal.Parse(KwSemiPunta),
                        KwKvarh = decimal.Parse(KwKvarh),
                        KwFp = decimal.Parse(KwFp),
                        Tipo = tipo,
                        Concepto = concepto,
                        NombreCliente = nomCliente,
                        Rpu = rpu
                    });
                }
            }
            return lst;
        }

    }
}
