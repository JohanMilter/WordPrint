class Program
{
    static void Main()
    {
        int lines = 1;
        bool wordPrint = true;
        string fileName = "WordTest1.docx";
        string filePath = @"D:\CodeProjects\GitHub\WordPrint\WordDocuments";

        if (wordPrint)
        {
            List<WordPrint.WordItem> wordItems = new();
            WordPrint.CreateDocument.CreateDoc(fileName, filePath);
            for (int i = 0; i < 1000; i++)
            {
                wordItems.Add(new WordPrint.WordItem
                {
                    Content = $"f({i})",
                    MathField = true,
                    NewLine = true,
                });
            }
            
            WordPrint.CreateDocument.EditDoc(Path.Combine(filePath, fileName), wordItems);
            foreach (string count in WordPrint.CalculateTime.CreateCounterList)
                Console.WriteLine(count);
            Console.WriteLine($"Total: {WordPrint.CalculateTime.CreateCount} ms");
            
        }
        else
        {
            Console.WriteLine($"It will take {WordPrint.CalculateTime.TimeForLines(lines)}ms to print {lines} line(s) in word document");
        }
    }
}

