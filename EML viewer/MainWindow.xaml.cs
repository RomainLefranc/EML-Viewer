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
using System.IO;
using Microsoft.Win32;
namespace EML_viewer
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

        private void click_filepicker(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EML files (*.eml)|*.eml";
            openFileDialog.InitialDirectory = @"C:\";
            if (openFileDialog.ShowDialog() == true) {
                

            }
        }

        private void source_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(@"c:\source\"))
            {
                Directory.CreateDirectory(@"c:\source\");
            }

            string[] filePaths = Directory.GetFiles(@"c:\source\", "*.eml");

            foreach (string file in filePaths)
            {
                source.Items.Add(file);
            }
        }

        private void source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FileInfo fileInfo = new FileInfo((source.SelectedItem).ToString());
            MsgReader.Mime.Message eml = MsgReader.Mime.Message.Load(fileInfo);
            from.Text = eml.Headers.From.ToString();
            if (eml.Headers != null)
            {
                if (eml.Headers.To != null)
                {
                    to.Text = string.Join(", ", eml.Headers.To.Select(x => x.Address).ToArray());

                }

            }
            subject.Text = eml.Headers.Subject;
            cc.Text = string.Join(", ", eml.Headers.Cc.Select(x => x.Address).ToArray());
            cci.Text = string.Join(", ", eml.Headers.Bcc.Select(x => x.Address).ToArray());
            date.Text = eml.Headers.Date;
            if (eml.TextBody != null)
            {
                htmlBody.Visibility = Visibility.Hidden;
                textBody.Visibility = Visibility.Visible;
                textBody.Text = System.Text.Encoding.UTF8.GetString(eml.TextBody.Body);
            }

            if (eml.HtmlBody != null)
            {
                htmlBody.Visibility = Visibility.Visible;
                textBody.Visibility = Visibility.Hidden;
                htmlBody.NavigateToString(System.Text.Encoding.UTF8.GetString(eml.HtmlBody.Body));
            }
        }
    }
}
