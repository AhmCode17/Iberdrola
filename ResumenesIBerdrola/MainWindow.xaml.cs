using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using System.Windows;
namespace ResumenesIBerdrola
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("hola");
        }

        private void btnSalir_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
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
    }
}
