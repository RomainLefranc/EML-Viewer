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
        private MsgReader.Mime.Message eml;
        public MainWindow()
        {
            InitializeComponent();
            // verification que le dossier source existe et creation du dossier si il n'existe pas
            if (!Directory.Exists(@"source\"))
            {
                _ = Directory.CreateDirectory(@"source\");
            }
        }

        private void Source_Loaded(object sender, RoutedEventArgs e)
        {

            // récupération d'une tableau contenant tout les fichiers du dossier source
            string[] filePaths = Directory.GetFiles(@"source\", "*.eml");

            // ajout des fichiers dans un combobox
            foreach (string file in filePaths)
            {
                _ = source.Items.Add(file.Split(@"\")[1]);
            }
        }

        private void Source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // récuperation du fichier a analyser
            FileInfo fileInfo = new FileInfo(@"source\" + source.SelectedItem.ToString());

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
                _ = attachment.Items.Add(piece.FileName + " - " + piece.ContentType.MediaType + " - " + piece.BodyEncoding.GetByteCount(piece.GetBodyAsText()) + " octets");
            }
            attachment.IsEnabled = attachment.Items.Count != 0;
            extract_attachment.IsEnabled = attachment.Items.Count != 0;

            textBody.Text = Encoding.UTF8.GetString(eml.TextBody.Body) + Encoding.UTF8.GetString(eml.HtmlBody.Body);
        }

        private void Extract_attachment_Click(object sender, RoutedEventArgs e)
        {
            if (attachment.SelectedIndex >= 0)
            {
                string filename = attachment.SelectedItem.ToString().Split(" - ")[0];
                if (!File.Exists(@"source\" + filename))
                {
                    MsgReader.Mime.MessagePart file = eml.Attachments.FirstOrDefault(x => x.FileName == filename);
                    FileInfo fileInfo = new FileInfo(@"source\" + filename);
                    file.Save(fileInfo);
                    _ = MessageBox.Show("Le fichier " + filename + " à été extrait");
                } else
                {
                    _ = MessageBox.Show("Le fichier " + filename + " à déjà été extrait");
                }

            } else
            {
                _ = MessageBox.Show("Veuillez selectioner une piece jointe");
            }
        }

        private void Openfolder_Click(object sender, RoutedEventArgs e)
        {
            _ = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = @"source\",
                UseShellExecute = true,
                Verb = "open"
            });
        }
    }
}
