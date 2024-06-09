using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;
using Spire.Xls;
using Microsoft.Win32;
using Spire.Xls.Core;
using System.IO;
using ImapX.Collections;
using ImapX;


namespace PW9
{
    /// <summary>
    /// Логика взаимодействия для ExelWindow.xaml
    /// </summary>
    public partial class ExelWindow : Window
    {
        public ExelWindow()
        {
            InitializeComponent();

            MinHeight = 450;
            MinWidth = 900;

            DataTable dt = new DataTable();

            dt.Columns.Add("");

            dgR.ItemsSource = dt.DefaultView;
        }

        private void CreateColumn_Click(object sender, RoutedEventArgs e)
        {
            Create_Column();
        }
        private void CreateRow_Click(object sender, RoutedEventArgs e)
        {
            Create_Row();
        }
        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            Delete_Row();
        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            Clear_DataGrid();
        }
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            Open_File();
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Save_File();
        }
        private void Email_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Email()
        {

        }
        private void Create_Column()
        {
            DataTable dataTable = (dgR.ItemsSource as DataView).Table;
            string name = tbX.Text;

            if (!dataTable.Columns.Contains(name))
            {
                dataTable.Columns.Add(name);
            }

            DataView dataView = dataTable.DefaultView;
            dgR.ItemsSource = null;      //четсно, не я писал это, а гпт, я в душе не понимаю почему и для чего это. Типо тут как-бы очищается датагрид, а для чего я непон люти словил
            dgR.ItemsSource = dataView;

        }
        private void Create_Row()
        {
            try
            {
                string row = "";
                dgR.Items.Add(row);
            }
            catch { }
        }

        private void Clear_DataGrid()
        {
            dgR.ItemsSource = null;
            DataTable dt = new DataTable();

            dt.Columns.Add("");

            dgR.ItemsSource = dt.DefaultView;
        }


        private void Open_File()
        {
            try
            {

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;

                    Workbook wb = new Workbook();
                    wb.LoadFromFile(filePath);

                    Worksheet sheet = wb.Worksheets[0];
                    CellRange locatedRange = sheet.AllocatedRange;

                    var dataTable = sheet.ExportDataTable(locatedRange, true);
                    dgR.ItemsSource = dataTable.DefaultView;
                }
            }
            catch { }
        }

        private void Save_File()
        {
            try
            {

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;

                    var dataTable = dgR.ItemsSource as DataView;

                    Workbook wb = new Workbook();
                    wb.Worksheets.Clear();
                    Worksheet sheet = wb.Worksheets.Add("лист 1");

                    sheet.InsertDataView(dataTable, true, 1, 1);

                    wb.SaveToFile(filePath, FileFormat.Version2016);
                }
            }
            catch { }
        }
        private void Delete_Row()
        {
            if (dgR.ItemsSource == null)
            {
                MessageBox.Show("Таблица пустая");
            }
            else
            {
                if (dgR.SelectedItem is DataRowView selectedRow)
                {
                    DataTable dataTable = ((DataView)dgR.ItemsSource).Table;
                    dataTable.Rows.Remove(selectedRow.Row);
                }
                else
                {
                    MessageBox.Show("Низя");
                }
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            EmailWindow window = new EmailWindow();
            window.Show();
        }
    }
}
