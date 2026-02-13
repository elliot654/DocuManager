using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Xml.Linq;
using System.Xml.Serialization;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


namespace DocuManager
{
    public partial class MainWindow : Window
    {
        private GridViewColumnHeader _lastHeaderClicked;
        private ListSortDirection _lastDirection;
        private Point _dragStartPoint;
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
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string file in files)
            {
                ProcessFile(file);
            }
        }
        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn)
                return;

            if (!Properties.Settings.Default.SkipRemove)
            {
                RemoveDialog dialog = new RemoveDialog
                {
                    Owner = this
                };

                if (dialog.ShowDialog() != true)
                    return;

                if (dialog.DontShowAgain)
                {
                    Properties.Settings.Default.SkipRemove = true;
                    Properties.Settings.Default.Save();
                }
            }

            Guid itemId = (Guid)btn.Tag;
            var match = records.FirstOrDefault(r => r.Id == itemId);

            if (match != null)
            {
                records.Remove(match);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Text & Documents|*.txt;*.rtf;*.docx;*.odt"
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    ProcessFile(file);
                }
            }
        }
        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is not GridViewColumnHeader header)
                return;

            if (header.Column.DisplayMemberBinding is not Binding binding)
                return;

            string sortBy = binding.Path.Path;

            ListSortDirection direction;

            if (header == _lastHeaderClicked)
            {
                direction = _lastDirection == ListSortDirection.Ascending
                    ? ListSortDirection.Descending
                    : ListSortDirection.Ascending;
            }
            else
            {
                direction = ListSortDirection.Ascending;
            }

            ICollectionView view = CollectionViewSource.GetDefaultView(ItemsList.ItemsSource);
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(sortBy, direction));
            view.Refresh();

            _lastHeaderClicked = header;
            _lastDirection = direction;
        }
        private void ItemsList_Drop(object sender, DragEventArgs e)
        {
            ICollectionView view = CollectionViewSource.GetDefaultView(records);

            if (view.SortDescriptions.Count > 0)
            {
                MessageBox.Show("Clear sorting before manually reordering.",
                                "Sorting Active",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                return;
            }

            if (!e.Data.GetDataPresent(typeof(Record)))
                return;

            Record droppedData = e.Data.GetData(typeof(Record)) as Record;
            Record target = ((FrameworkElement)e.OriginalSource).DataContext as Record;

            if (droppedData == null || target == null || droppedData == target)
                return;

            int removedIdx = records.IndexOf(droppedData);
            int targetIdx = records.IndexOf(target);

            records.Move(removedIdx, targetIdx);
        }
        private void ItemsList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
        }
        private void ItemsList_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            Point position = e.GetPosition(null);

            if (Math.Abs(position.X - _dragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(position.Y - _dragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            if (ItemsList.SelectedItem == null)
                return;

            DragDrop.DoDragDrop(ItemsList, ItemsList.SelectedItem, DragDropEffects.Move);
        }
        private void ProcessFile(string filePath)
        {
            string filename = Path.GetFileName(filePath);
            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            string plainText = null;
            int paragraphs = 0;

            if (extension == ".txt" || extension == ".rtf")
            {
                string text = File.ReadAllText(filePath);

                paragraphs = Regex
                    .Split(text.Trim(), @"(\r?\n\s*\r?\n)+")
                    .Count(p => !string.IsNullOrWhiteSpace(p));

                // Use RichTextBox to normalize RTF → plain text
                RichTextBox richTextBox = new RichTextBox();
                TextRange range = new TextRange(
                    richTextBox.Document.ContentStart,
                    richTextBox.Document.ContentEnd);

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
                {
                    range.Load(stream, DataFormats.Rtf);
                }

                plainText = range.Text;
            }
            else if (extension == ".docx")
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    StringBuilder sb = new StringBuilder();
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        sb.AppendLine(para.InnerText);
                        paragraphs++;
                    }

                    plainText = sb.ToString();
                }
            }
            else if (extension == ".odt")
            {
                using (ZipArchive archive = ZipFile.OpenRead(filePath))
                {
                    ZipArchiveEntry entry = archive.GetEntry("content.xml");
                    using StreamReader reader = new StreamReader(entry.Open());

                    var xml = XDocument.Parse(reader.ReadToEnd());
                    XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

                    StringBuilder sb = new StringBuilder();
                    foreach (var para in xml.Descendants(textNs + "p"))
                    {
                        sb.AppendLine(para.Value);
                        paragraphs++;
                    }

                    plainText = sb.ToString();
                }
            }

            if (string.IsNullOrWhiteSpace(plainText))
                return;

            int characters = plainText.Length;
            int words = Regex.Matches(plainText, @"\b\w+\b").Count;

            Guid id = Guid.NewGuid();
            records.Add(new Record(id, filename, characters, words, paragraphs, extension));
        }
    }
}