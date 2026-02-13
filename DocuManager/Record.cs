using System.ComponentModel;

namespace DocuManager
{
[Serializable]
public class Record : INotifyPropertyChanged
{
    public Record() { } // Parameterless constructor required by XmlSerializer

    public Record(Guid id, string title, int characters, int words, int paragraphs, string extension, string filePath)
    {
        Id = id;
        Title = title;
        Characters = characters;
        Words = words;
        Paragraphs = paragraphs;
        Extension = extension;
        FilePath = filePath;
    }

    public Guid Id { get; set; }
    public string Title { get; set; }
    public string Extension { get; set; }
    public string FilePath { get; set; }

    private int _characters;
    public int Characters
    {
        get => _characters;
        set { _characters = value; OnPropertyChanged(nameof(Characters)); }
    }

    private int _words;
    public int Words
    {
        get => _words;
        set { _words = value; OnPropertyChanged(nameof(Words)); }
    }

    private int _paragraphs;
    public int Paragraphs
    {
        get => _paragraphs;
        set { _paragraphs = value; OnPropertyChanged(nameof(Paragraphs)); }
    }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
}
