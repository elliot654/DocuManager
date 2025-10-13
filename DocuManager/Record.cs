namespace DocuManager
{
    public class Record
    {
        public Guid Id { get; set; }
        public string Title { get; set; }
        public int Characters { get; set; }
        public int Words { get; set; }
        public int Paragraphs { get; set; }
        public string Extension { get; set; }

        public Record() {
            Title = "";
            Characters = 0;
            Words = 0;
            Paragraphs = 0;            
            Extension = "";
        }

        public Record(Guid id, string title, int characters, int words, int paragraphs, string extension)
        {
            this.Id = id;
            this.Title = title;
            this.Characters = characters;
            this.Words = words;
            this.Paragraphs = paragraphs;           
            this.Extension = extension;
        }

    }
}
