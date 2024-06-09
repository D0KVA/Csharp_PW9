using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
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

namespace PW9
{
    /// <summary>
    /// Логика взаимодействия для EmailWindow.xaml
    /// </summary>
    public partial class EmailWindow : Window
    {
        public EmailWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            TextRange range = new TextRange(MessageRtb.Document.ContentStart, MessageRtb.Document.ContentEnd);

            MailMessage message = new MailMessage(From.Text, To.Text, Subject.Text, range.Text);
            message.IsBodyHtml = true;


            SmtpClient client = new SmtpClient("smtp.mail.ru");
            client.Credentials = new NetworkCredential(From.Text, Pass.Text);
            client.EnableSsl = true;
            client.Send(message);
        }
    }
}
