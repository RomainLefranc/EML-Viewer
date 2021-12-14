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
        MsgReader.Mime.Message eml;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Source_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(@"c:\source\"))
            {
                _ = Directory.CreateDirectory(@"c:\source\");
            }

            string[] filePaths = Directory.GetFiles(@"c:\source\", "*.eml");

            foreach (string file in filePaths)
            {
                _ = source.Items.Add(file);
            }
        }

        private void Source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FileInfo fileInfo = new FileInfo(source.SelectedItem.ToString());
            eml = MsgReader.Mime.Message.Load(fileInfo);
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

            attachment.Items.Clear();
            foreach (MsgReader.Mime.MessagePart piece in eml.Attachments)
            {
                _ = attachment.Items.Add(piece.ContentType + " - " + piece.BodyEncoding.GetByteCount(piece.GetBodyAsText()) + " octets");
            }

            if (attachment.Items.Count == 0)
            {
                attachment.IsEnabled = false;
            } else
            {
                attachment.IsEnabled = true;
            }

            if (eml.TextBody != null)
            {
                textBody.Text = Encoding.UTF8.GetString(eml.TextBody.Body);
            }
            if (eml.HtmlBody != null)
            {
                textBody.Text = Encoding.UTF8.GetString(eml.HtmlBody.Body);
            }
        }

    }
}
