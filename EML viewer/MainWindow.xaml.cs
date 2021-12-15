using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.IO;
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
            // verification que le dossier source existe et creation du dossier si il n'existe pas
            if (!Directory.Exists(@"c:\source\"))
            {
                _ = Directory.CreateDirectory(@"c:\source\");
            }

            // récupération d'une tableau contenant tout les fichiers du dossier source
            string[] filePaths = Directory.GetFiles(@"c:\source\", "*.eml");

            // ajout des fichiers dans un combobox
            foreach (string file in filePaths)
            {
                _ = source.Items.Add(file);
            }
        }

        private void Source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // récuperation du fichier a analyser
            FileInfo fileInfo = new FileInfo(source.SelectedItem.ToString());

            // analyse du fichier
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

            attachment.IsEnabled = attachment.Items.Count != 0;

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
