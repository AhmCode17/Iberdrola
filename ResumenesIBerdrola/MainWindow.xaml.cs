using log4net;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using ResumenesIBerdrola.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ResumenesIBerdrola
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private ErrorModel errorModel = new ErrorModel();
        MsAccessDataContext db;
        bool reemplazarData = false;
        private string PathDataBase = string.Empty;
        string[] tipos = { "ENERGIA TOTAL", "ENERGIA NORMAL", "ENERGIA PORTEADA", "ENERGIA NORMAL POR FALTANTE", "RESPALDO POR CARGA" };
        public MainWindow()
        {
            InitializeComponent();
           
            try
            {
                var appenders = log4net.LogManager.GetRepository().GetAppenders();
                foreach (var appender in appenders)
                {
                    var rollingFileAppender = appender as log4net.Appender.RollingFileAppender;
                    if (rollingFileAppender != null)
                    {
                        rollingFileAppender.ImmediateFlush = true;
                        rollingFileAppender.LockingModel = new log4net.Appender.FileAppender.MinimalLock();
                        rollingFileAppender.ActivateOptions();
                    }
                }
                //  File.Delete(@"C:\Iberdrola\iberdrolaLog.log");
                string logPath = @"C:\Iberdrola\iberdrolaLog.log";
                if (File.Exists(logPath))
                {

                    FileInfo fi = new FileInfo(logPath);

                    var logFiles = fi.Directory.GetFiles(fi.Name + "*");

                    foreach (var log in logFiles)
                    {
                        if (File.Exists(log.FullName)) File.Delete(log.FullName);
                    }
                }


            }
            catch (Exception ex)
            {
            }
            log4net.Config.XmlConfigurator.Configure();
            btnProcesar.IsEnabled = false;
        }

        public IDbConnection Connection { get; }

        public List<ConceptoModel> Conceptos = new List<ConceptoModel>();
        public List<CentralModel> Centrales = new List<CentralModel>();
        public List<string> FilesExcelNew = new List<string>();
        public List<string> FilesExcelOld = new List<string>();

        private void BtnSalir_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnSeleccionar_Click(object sender, RoutedEventArgs e)
        {
            lstFiles.Items.Clear();
            FilesExcelNew = new List<string>();
            FilesExcelOld = new List<string>();
            lblFilesSuccess.Text = "";

            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                InitialDirectory = @"C:\",
                IsFolderPicker = true
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtRuta.Text = dialog.FileName;

                foreach (string file in Directory.EnumerateFiles(txtRuta.Text, "*.xlsx"))
                {
                    lstFiles.Items.Add(file);
                    FilesExcelNew.Add(file);
                }

                foreach (string file in Directory.EnumerateFiles(txtRuta.Text, "*.xls"))
                {
                    lstFiles.Items.Add(file);
                    FilesExcelOld.Add(file);
                    lblFiles.Content = lstFiles.Items.Count + " archivos";
                }
            }
        }

        private void BtnProcesar_Click(object sender, RoutedEventArgs e)
        {
            if (lstFiles.Items.Count >= 1)
            {
                var task = new Thread(new ThreadStart(delegate
                {
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action<ProgressBar>(SetControls), pbStatus);
                    ExecuteMigration();
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action<ProgressBar>(RemoveControls), pbStatus);
                }));
                task.Start();
            }
            else
            {
                MessageBox.Show("No hay archivos para procesar", "Iberdrola", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SetControls(ProgressBar progressBar)
        {
            reemplazarData = bool.Parse(chkReemplazar.IsChecked.ToString());
            btnSeleccionarBd.IsEnabled = false;
            btnLog.IsEnabled = false;
            btnProcesar.IsEnabled = false;
            chkReemplazar.IsEnabled = false;
            btnSalir.IsEnabled = false;
            btnSeleccionar.IsEnabled = false;
            progressBar.Visibility = Visibility.Visible;
        }

        private void RemoveControls(ProgressBar progressBar)
        {
            progressBar.Visibility = Visibility.Hidden;
            btnSeleccionarBd.IsEnabled = true;
            btnLog.IsEnabled = true;
            btnProcesar.IsEnabled = true;
            btnSalir.IsEnabled = true;
            chkReemplazar.IsEnabled = true;
            btnSeleccionar.IsEnabled = true;
        }
        public void ExecuteMigration()
        {
            Conceptos = (List<ConceptoModel>)db.GetConcepto().Data;
            Centrales = (List<CentralModel>)db.GetCentral().Data;

            var suma = 50 / FilesExcelNew.Count();
            var i = 0;
            foreach (var item in FilesExcelNew)
            {
                GetHeaderExcel(item);

                i++;
                //lblFilesSuccess.Text = i + " archivos procesados";
            }
            //i = i + (50 / FilesExcelOld.Count());
            foreach (var item in FilesExcelOld)
            {
                GetHeaderExcelOld(item);

                i += i;
                // lblFilesSuccess.Text = i + " archivos procesados";
            }
            MessageBox.Show("Se terminó el proceso favor de revisar el log", "Iberdrola", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }


        public void GetHeaderExcel(string path)
        {
            try
            {
                log.Info(string.Format("*************** Se va a leer el archivo: {0} ***************", path));
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                //FileInfo existingFile = new FileInfo(@"C:\Users\Babel\Documents\Iberdrola\test\Resumen BNS.xlsx");
                FileInfo existingFile = new FileInfo(path);
                List<ResumenModel> lstresumenModelHeader = new List<ResumenModel>();
                List<ResumenModel> lstDetail = new List<ResumenModel>();
                ResumenBaseModel resumen = new ResumenBaseModel();
                int fkCentral = 0;
                bool cont = true;
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    //get the first worksheet in the workbook
                    var currentSheet = package.Workbook.Worksheets;
                    ExcelWorksheet worksheet = currentSheet.First(x => x.Hidden != eWorkSheetHidden.Hidden);
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    var periodoText = worksheet.Cells[4, 1].Value.ToString();
                    var perido = worksheet.Cells[4, 1].Value.ToString().Split(':')[0];
                    var central = worksheet.Cells[3, 1].Value.ToString().Split(':')[1].Trim();
                    var tipo = worksheet.Cells[17, 2].Value.ToString();
                    int endLetter = perido.Length - 7;
                    perido = perido.Substring(0, endLetter);

                    int row2 = 17;
                    var centralFind = Centrales.FirstOrDefault(x => x.Cliente.Contains(central));
                    if (centralFind != null)
                    {
                        fkCentral = centralFind.Id;

                        var data = db.SaveResumen(new ResumenBaseModel
                        {
                            FkCentral = fkCentral,
                            Periodo = perido,
                            FechaCreacion = DateTime.Now,
                            Central = central,
                            Reemplazar = reemplazarData
                        });
                        if (data.Success)
                        {
                            var resumenList = (ResumenBaseModel)data.Data;
                            resumen.Id = resumenList.Id;
                        }
                        else
                        {
                            cont = false;
                        }
                    }
                    if (cont)
                    {
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
                            int fkConcepto = 0;

                            var conceptoFind = Conceptos.FirstOrDefault(x => x.Concepto.Contains(concepto));
                            if (conceptoFind != null)
                                fkConcepto = conceptoFind.Id;

                            lstresumenModelHeader.Add(new ResumenModel
                            {
                                KwhBase = GetValue(KwhBase),
                                KwhIntermedia = GetValue(KwhIntermedia),
                                KwhPunta = GetValue(KwhPunta),
                                KwhSemiPunta = GetValue(KwhSemiPunta),
                                KwhTotales = GetValue(KwhTotales),
                                KwBase = GetValue(KwBase),
                                KwIntermedia = GetValue(KwIntermedia),
                                KwPunta = GetValue(KwPunta),
                                KwSemiPunta = GetValue(KwSemiPunta),
                                KwKvarh = GetValue(KwKvarh),
                                KwFp = GetValue(KwFp),
                                Concepto = concepto,
                                Periodo = perido,
                                FkResumen = resumen.Id,
                                FkConcepto = fkConcepto,
                                Reemplazar = reemplazarData
                            });
                        }

                        //Guarda el header en la tabla de 
                        foreach (var item in lstresumenModelHeader)
                        {
                            var resp = db.SaveCentralTotal(item);
                        }

                        //Obtiene el detalle
                        lstDetail.AddRange(GetDetail(worksheet, row2, rowCount, tipos, tipo));
                        foreach (var item in lstDetail)
                        {
                            item.FkResumen = resumen.Id;
                            var resp = db.SavePuntoDeCarga(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(string.Format("No es un archivo con formato válido: {0}", ex.Message));
            }
            log.Info(string.Format("*************** Fin del archivo archivo: {0} ***************", path));
        }

        public decimal GetValue(string vl)
        {
            decimal data;
            decimal.TryParse(vl, out data);
            return data;
        }
        public List<ResumenModel> GetDetail(ExcelWorksheet worksheet, int row2, int rowCount, string[] tipos, string tipo)
        {
            List<ResumenModel> lst = new List<ResumenModel>();
            try
            {
                for (int row = row2; row <= rowCount; row++)
                {
                    var baseTipo = worksheet.Cells[row, 2].Value == null ? string.Empty : worksheet.Cells[row, 2].Value.ToString();
                    if (baseTipo == string.Empty)
                        break;
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

                        var capacidadTotal = worksheet.Cells[row, 39].Value == null ? string.Empty : worksheet.Cells[row, 39].Value.ToString();
                        var distribucionTotal = worksheet.Cells[row, 40].Value == null ? string.Empty : worksheet.Cells[row, 40].Value.ToString();

                        var nomCliente = concepto.Split('_')[2];
                        var rpu = concepto.Split('_')[1];


                        int fkConcepto = 0;

                        var conceptoFind = Conceptos.FirstOrDefault(x => x.Concepto.Contains(tipo));
                        if (conceptoFind != null)
                            fkConcepto = conceptoFind.Id;
                        lst.Add(new ResumenModel
                        {
                            KwhBase = GetValue(KwhBase),
                            KwhIntermedia = GetValue(KwhIntermedia),
                            KwhPunta = GetValue(KwhPunta),
                            KwhSemiPunta = GetValue(KwhSemiPunta),
                            KwhTotales = GetValue(KwhTotales),
                            KwBase = GetValue(KwBase),
                            KwIntermedia = GetValue(KwIntermedia),
                            KwPunta = GetValue(KwPunta),
                            KwSemiPunta = GetValue(KwSemiPunta),
                            KwKvarh = GetValue(KwKvarh),
                            KwFp = GetValue(KwFp),
                            CapacidadTotal = GetValue(capacidadTotal),
                            DistribucionTotal = GetValue(distribucionTotal),
                            Tipo = tipo,
                            Descripcion = concepto,
                            NombreCliente = nomCliente,
                            Rpu = rpu,
                            FkConcepto = fkConcepto,
                            Reemplazar = reemplazarData
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(string.Format("Ocurrio un error: {0}", ex.Message));
            }
            return lst;
        }

        public void GetHeaderExcelOld(string path)
        {
            try
            {
                log.Info(string.Format("*************** Se va a leer el archivo: {0} ***************", path));
                DataTable dtTable = new DataTable();
                List<string> rowList = new List<string>();
                ISheet sheet;
                List<ResumenModel> lst = new List<ResumenModel>();
                List<ResumenModel> lstDetail = new List<ResumenModel>();
                ResumenBaseModel resumen = new ResumenBaseModel();
                int fkCentral = 0;
                bool cont = true;
                //using (var stream = new FileStream(@"C:\Users\Babel\Documents\Iberdrola\test\Resumen CDUII.xls", FileMode.Open))
                using (var stream = new FileStream(path, FileMode.Open))
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
                        central = cellCentral.ToString().Split(':')[1].Trim();

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

                    var centralFind = Centrales.FirstOrDefault(x => x.Cliente.Contains(central));
                    if (centralFind != null)
                    {
                        fkCentral = centralFind.Id;
                        var data = db.SaveResumen(new ResumenBaseModel
                        {
                            FkCentral = fkCentral,
                            Periodo = perido,
                            FechaCreacion = DateTime.Now,
                            Central = central,
                            Reemplazar = reemplazarData
                        });
                        if (data.Success)
                        {
                            var resumenList = (ResumenBaseModel)data.Data;
                            resumen.Id = resumenList.Id;
                        }
                        else
                        {
                            cont = false;

                        }
                    }
                    if (cont)
                    {
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

                            int fkConcepto = 0;

                            var conceptoFind = Conceptos.FirstOrDefault(x => x.Concepto.Contains(concepto));
                            if (conceptoFind != null)
                                fkConcepto = conceptoFind.Id;

                            lst.Add(new ResumenModel
                            {
                                KwhBase = GetValue(KwhBase),
                                KwhIntermedia = GetValue(KwhIntermedia),
                                KwhPunta = GetValue(KwhPunta),
                                KwhSemiPunta = GetValue(KwhSemiPunta),
                                KwhTotales = GetValue(KwhTotales),
                                KwBase = GetValue(KwBase),
                                KwIntermedia = GetValue(KwIntermedia),
                                KwPunta = GetValue(KwPunta),
                                KwSemiPunta = GetValue(KwSemiPunta),
                                KwKvarh = GetValue(KwKvarh),
                                KwFp = GetValue(KwFp),
                                Concepto = concepto,
                                Periodo = perido,
                                FkResumen = resumen.Id,
                                FkConcepto = fkConcepto,
                                Reemplazar = reemplazarData
                            });
                        }

                        //Guarda el header en la tabla de 
                        foreach (var item in lst)
                        {
                            var resp = db.SaveCentralTotal(item);
                        }

                        //Obtiene el detalle
                        lstDetail.AddRange(GetDetailOld(sheet, 16, sheet.LastRowNum, tipos, tipo));

                        foreach (var item in lstDetail)
                        {
                            item.FkResumen = resumen.Id;
                            var resp = db.SavePuntoDeCarga(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(string.Format("No es un archivo con formato válido: {0}", ex.Message));
            }
            log.Info(string.Format("*************** Fin del archivo archivo: {0} ***************", path));
        }

        public List<ResumenModel> GetDetailOld(ISheet sheet, int row2, int rowCount, string[] tipos, string tipo)
        {
            List<ResumenModel> lst = new List<ResumenModel>();
            try
            {
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
                        var concepto = worksheet.GetCell(1) == null ? string.Empty : worksheet.GetCell(1).ToString();
                        var KwhBase = worksheet.GetCell(2) == null ? string.Empty : worksheet.GetCell(2).NumericCellValue.ToString();
                        var KwhIntermedia = worksheet.GetCell(3) == null ? string.Empty : worksheet.GetCell(3).NumericCellValue.ToString();
                        var KwhPunta = worksheet.GetCell(4) == null ? string.Empty : worksheet.GetCell(4).NumericCellValue.ToString();
                        var KwhSemiPunta = worksheet.GetCell(5) == null ? string.Empty : worksheet.GetCell(5).NumericCellValue.ToString();
                        var KwhTotales = worksheet.GetCell(6) == null ? string.Empty : worksheet.GetCell(6).NumericCellValue.ToString();
                        var KwBase = worksheet.GetCell(7) == null ? string.Empty : worksheet.GetCell(7).NumericCellValue.ToString();
                        var KwIntermedia = worksheet.GetCell(8) == null ? string.Empty : worksheet.GetCell(8).NumericCellValue.ToString();
                        var KwPunta = worksheet.GetCell(9) == null ? string.Empty : worksheet.GetCell(9).NumericCellValue.ToString();
                        var KwSemiPunta = worksheet.GetCell(10) == null ? string.Empty : worksheet.GetCell(10).NumericCellValue.ToString();
                        var KwKvarh = worksheet.GetCell(11) == null ? string.Empty : worksheet.GetCell(11).NumericCellValue.ToString();
                        var KwFp = worksheet.GetCell(12) == null ? string.Empty : worksheet.GetCell(12).NumericCellValue.ToString();
                        var nomCliente = concepto.Split('_')[2];
                        var rpu = concepto.Split('_')[1];

                        var capacidadTotal = worksheet.GetCell(38) == null ? string.Empty : worksheet.GetCell(38).NumericCellValue.ToString();
                        var distribucionTotal = worksheet.GetCell(39) == null ? string.Empty : worksheet.GetCell(39).NumericCellValue.ToString();

                        int fkConcepto = 0;

                        var conceptoFind = Conceptos.FirstOrDefault(x => x.Concepto.Contains(tipo));
                        if (conceptoFind != null)
                            fkConcepto = conceptoFind.Id;
                        lst.Add(new ResumenModel
                        {
                            KwhBase = GetValue(KwhBase),
                            KwhIntermedia = GetValue(KwhIntermedia),
                            KwhPunta = GetValue(KwhPunta),
                            KwhSemiPunta = GetValue(KwhSemiPunta),
                            KwhTotales = GetValue(KwhTotales),
                            KwBase = GetValue(KwBase),
                            KwIntermedia = GetValue(KwIntermedia),
                            KwPunta = GetValue(KwPunta),
                            KwSemiPunta = GetValue(KwSemiPunta),
                            KwKvarh = GetValue(KwKvarh),
                            KwFp = GetValue(KwFp),
                            CapacidadTotal = GetValue(capacidadTotal),
                            DistribucionTotal = GetValue(distribucionTotal),
                            Tipo = tipo,
                            Descripcion = concepto,
                            NombreCliente = nomCliente,
                            Rpu = rpu,
                            FkConcepto = fkConcepto,
                            Reemplazar = reemplazarData
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(string.Format("Ocurrio un error detail: {0}", ex.Message));
            }
            return lst;
        }

        private void btnLog_Click(object sender, RoutedEventArgs e)
        {
            // Read the file as one string.
            var scr = new ErrorsScr();
            scr.ShowDialog();
        }

        private void btnSeleccionarBd_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Access files (*.accdb)|*.accdb|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {                
                db = new MsAccessDataContext(log, openFileDialog.FileName);
                btnProcesar.IsEnabled = true;
            }
                
        }
    }
}
