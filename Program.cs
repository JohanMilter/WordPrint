class Program
{
    static void Main()
    {
        int lines = 5;
        bool wordPrint = true;
        string fileName = "WordTest1.docx";
        string filePath = @"D:\CodeProjects\GitHub\WordPrint\WordDocuments";

        if (wordPrint)
        {
            List<WordPrint.WordItem> wordItems = new();
            WordPrint.CreateDocument.CreateDoc(fileName, filePath);
            for (int i_2 = 0; i_2 < lines; i_2++)
            {
                wordItems.Add(new WordPrint.WordItem
                {
                    Content = $"f({i_2})",
                    MathField = false,
                    NewLine = false,
                });
                wordItems.Add(new WordPrint.WordItem
                {
                    Content = $"f({i_2})",
                    MathField = false,
                    NewLine = false,
                });
            }
            WordPrint.CreateDocument.EditDoc(Path.Combine(filePath, fileName), wordItems);
        }
        else
        {
            Console.WriteLine($"It will take {WordPrint.CalculateTime.TimeForLines(lines)}ms to print {lines} line(s) in word document");
        }
    }
}

