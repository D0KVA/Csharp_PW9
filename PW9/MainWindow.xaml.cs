using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PW9
{
    
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            MinHeight = 500;
            MinWidth = 450;
        }

        private void CreateExel_Click(object sender, RoutedEventArgs e)
        {
            ExelWindow window = new ExelWindow();
            window.Show();
            this.Close();
        }
    }
}