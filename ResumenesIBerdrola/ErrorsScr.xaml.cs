using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ResumenesIBerdrola
{
    /// <summary>
    /// Interaction logic for ErrorsScr.xaml
    /// </summary>
    public partial class ErrorsScr : Window
    {
        public ErrorsScr()
        {
            InitializeComponent();
            try
            {
                File.Copy(@"C:\Iberdrola\iberdrolaLog.log", @"C:\Iberdrola\iberdrolaLog2.log", true);
                string[] text = System.IO.File.ReadAllLines(@"C:\Iberdrola\iberdrolaLog2.log");
                //System.Console.WriteLine("Contents of WriteLines2.txt = ");
                foreach (string line in text)
                {
                    // Use a tab to indent each line of the file.
                    lstErrores.Items.Add("\t" + line);
                    //Console.WriteLine("\t" + line);
                }

                File.Delete(@"C:\Iberdrola\iberdrolaLog2.log");
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se encuentra el archivo de log", "Iberdrola", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
    }
}
