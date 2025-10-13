using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml.Linq;
using System.Xml.Serialization;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


namespace DocuManager
{
    public partial class MainWindow : Window
    {
        ObservableCollection<Record> records = new ObservableCollection<Record>();
        XmlSerializer serializer = new XmlSerializer(typeof(ObservableCollection<Record>));
        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;
            if (File.Exists("usrData.xml"))
            {
                using (StreamReader reader = new StreamReader("usrData.xml"))
                {
                    records = (ObservableCollection<Record>)serializer.Deserialize(reader);
                }
            }
            ItemsList.ItemsSource = records;
            records.CollectionChanged += (_, __) => UpdateHint();
            UpdateHint();
            
        }
        private void UpdateHint()
        {
            if (records.Count == 0)
            {
                HintLabel.Visibility = Visibility.Visible;
            }
            else
            {
                HintLabel.Visibility = Visibility.Hidden;
            }
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            StreamWriter streamWriter = new StreamWriter("usrData.xml");
            serializer.Serialize(streamWriter, records);
            streamWriter.Close();
        }

        private void Panel_Drop(object sender, DragEventArgs e)
        {
            if(e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string filename = Path.GetFileName(file);
                    string filePath = Path.GetFullPath(file);
                    string extension = Path.GetExtension(file);
                    string plainText = null;
                    int paragraphs = 0;
                    if (extension == ".txt" || extension == ".rtf")
                    {
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            string text = reader.ReadToEnd();
                            paragraphs = Regex.Split(text.Trim(), @"(\r?\n\s*\r?\n)+").Count(p => !string.IsNullOrWhiteSpace(p));
                            //this block sets the data from the file to the content of a textbox object to use windows built in controls to strip the RTF formatting, it doesnt affect txt files so they can be processed together here
                            RichTextBox richTextBox = new RichTextBox();
                            TextRange range = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                            using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(text)))
                            {
                                range.Load(stream, DataFormats.Rtf);
                            }
                            plainText = range.Text;

                        }
                    }
                    if (extension == ".docx")
                    {
                        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                        {
                            StringBuilder stringBuilder = new StringBuilder();
                            Body body = wordDoc.MainDocumentPart.Document.Body;
                            foreach (var para in body.Elements<Paragraph>())
                            {
                                stringBuilder.AppendLine(para.InnerText);
                                paragraphs++;
                            }
                            plainText = stringBuilder.ToString();
                        }
                    }
                    if (extension == ".odt")
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(filePath))
                        {
                            ZipArchiveEntry contentXmlEntry = archive.GetEntry("content.xml");
                            StreamReader reader = new StreamReader(contentXmlEntry.Open());
                            var xml = XDocument.Parse(reader.ReadToEnd());
                            XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
                            StringBuilder stringBuilder = new StringBuilder();
                            foreach (var para in xml.Descendants(textNs + "p"))
                            {
                                stringBuilder.AppendLine(para.Value);
                                paragraphs++;
                            }
                            plainText = stringBuilder.ToString();
                        }
                    }
                    int characters = plainText.Length;
                    int words = Regex.Matches(plainText, @"\b\w+\b").Count;
                    Guid id = Guid.NewGuid();
                    records.Add(new Record(id, filename, characters, words, paragraphs, extension));
                    
                }
            }
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                Guid item = (Guid)btn.Tag;
                var match = records.FirstOrDefault(record => record.Id == item);
                if (match != null)
                {
                    records.Remove(match);
                }
            }
           
        }
    }
}