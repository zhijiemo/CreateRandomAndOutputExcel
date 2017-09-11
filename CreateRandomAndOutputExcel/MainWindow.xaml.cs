using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System.IO;

namespace CreateRandomAndOutputExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        //SLDocument tl = new SLDocument("Miscellaneous.xlsx", "Sheet");
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateRandomTable_Click(object sender, RoutedEventArgs e)
        {
            new CreateRandomTable().createRandomTable();
        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            new CalculatorAndSetPropertyAndOutput().Sum();
        }

        private void Calculate1_Click(object sender, RoutedEventArgs e)
        {
            new CalculatorAndSetPropertyAndOutput().Subtract();
        }

        private void Calculate2_Click(object sender, RoutedEventArgs e)
        {
            new CalculatorAndSetPropertyAndOutput().Average();
        }

        private void CreateExcel_Click(object sender, RoutedEventArgs e)
        {
            new CalculatorAndSetPropertyAndOutput().SetPropertyAndOutput();
            new CalculatorAndSetPropertyAndOutput().SetCellDocument();
            new CalculatorAndSetPropertyAndOutput().GetPresentTime();
        }
    }
}
