﻿using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using Microsoft.Win32;
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
        private void Combobox_source_load()
        {
            // vide le select de fichier
            source.Items.Clear();
            // récupération d'une tableau contenant tout les fichiers du dossier source
            string[] filePaths = Directory.GetFiles(@"source\", "*.eml");

            if (filePaths.Length == 0)
            {
                source.IsEnabled = false;
            }
            else
            {
                source.IsEnabled = true;
                // ajout des fichiers dans un combobox
                foreach (string file in filePaths)
                {
                    _ = source.Items.Add(file.Split(@"\")[1]);
                }
            }
        }

        private void Source_Loaded(object sender, RoutedEventArgs e)
        {
            Combobox_source_load();
        }

        private void Source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // récuperation du nom du fichier a analyser
            FileInfo fileInfo = new FileInfo(@"source\" + source.SelectedItem.ToString());

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

            // vide la liste des pieces jointes
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
                Ookii.Dialogs.Wpf.VistaFolderBrowserDialog dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
                if (dialog.ShowDialog(this).GetValueOrDefault())
                {
                    MsgReader.Mime.MessagePart file = eml.Attachments.FirstOrDefault(x => x.FileName == filename);
                    FileInfo fileInfo = new FileInfo(dialog.SelectedPath + @"\" + filename);
                    file.Save(fileInfo);
                    _ = MessageBox.Show("Le fichier " + filename + " à été extrait");
                }
            }
            else
            {
                _ = MessageBox.Show("Veuillez selectionner une piece jointe");
            }
        }

        private void AddFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Eml files (*.eml)|*.eml"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    File.Copy(openFileDialog.FileName, @"source\" + Path.GetFileName(openFileDialog.FileName));
                }
                catch (IOException error)
                {
                    _ = MessageBox.Show(error.Message);
                }
            }
            Combobox_source_load();
        }
    }
}
