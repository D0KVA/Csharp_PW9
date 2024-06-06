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
using System.Windows.Shapes;
using static MaterialDesignThemes.Wpf.Theme;

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
        }

        private void CreateColumn_Click(object sender, RoutedEventArgs e)
        {
            DataGridTextColumn newColumn = new DataGridTextColumn();
            newColumn.Header = tbX.Text;
 
            dgR.Columns.Add(newColumn);
        }
        private void DeleteColumn_Click(object sender, RoutedEventArgs e)
        {

        }
        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Email_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
