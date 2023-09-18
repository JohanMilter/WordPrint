using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using W = DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

enum ParPos
{
    Left,
    Center,
    Right,
}
enum WordItemType
{
    Table,
    Text,
    Image,
}
enum TextUnderLine
{
    None,
    Single,
    Doulble,
    Dashed,
    Thick,
    LongDashed,
    DashLine,
    DashDashLine,
    Wave,
}
readonly struct ParInden
{
    public ParInden(int left, int right, int top, int bottom)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
    }
    public int Left { get; }
    public int Right { get; }
    public int Top { get; }
    public int Bottom { get; }
}
class WordItem
{
    //Text Options
    public string Text = "";
    public string Font = "";
    public int FontSize = 11;
    //Format Text
    public string BackgroundColor = "";
    public string FontColor = "";
    public bool Bold = false;
    public bool Italic = false;
    public TextUnderLine Underline = TextUnderLine.None;
    //Type
    public WordItemType ItemType = WordItemType.Text;
    //Math
    public bool MathField = false;
    public bool XML = false;
    //Position
    public ParPos Position = ParPos.Left;
    public ParInden Margin = new(0,0,0,0);
}
class CreateDocument
{
    public string? FilePath;
    public static void CreateDoc(string DocumentName, string desktopPath)
    {
        try
        {
            if (!Directory.Exists(desktopPath))
                Directory.CreateDirectory(desktopPath);
            #region Create Word-Document
            using WordprocessingDocument wordDoc = WordprocessingDocument.Create(Path.Combine(desktopPath, $"{DocumentName}"), WordprocessingDocumentType.Document);
            MainDocumentPart MainPart = wordDoc.AddMainDocumentPart();
            MainPart.Document = new W.Document
            {
                Body = new W.Body()
            };
            MainPart.Document.Save();
            #endregion
        }
        catch
        {
            Console.WriteLine("Dokumentet er allerede åben. Luk word dokumentet!");
        }
    }

    /// <summary>
    /// <br>Creates the body to append to main body</br>
    /// </summary>
    /// <param name="collection"></param>
    public static W.Body CreateBody(List<WordItem> collection)
    {
        W.Body body = new();
        List<(WordItemType, List<WordItem>)> itemCollections = new();
        List<WordItem> currentItems = new();
        for (int i = 0; i < collection.Count; i++)
        {
            currentItems.Add(collection[i]);
            if ((i + 1) == collection.Count || ((i + 1) < collection.Count && collection[i].ItemType != collection[i + 1].ItemType))
            {
                itemCollections.Add((collection[i].ItemType, currentItems));
                currentItems = new();
            }
        }
        foreach ((WordItemType type, List<WordItem> items) collect in itemCollections)
        {
            if (collect.type == WordItemType.Table)
            {
                body.AppendChild(CreateWordItemStyles.CreateTable(collect.items));
            }
            else if (collect.type == WordItemType.Text)
            {
                foreach (WordItem wordItem in collect.items)
                    body.AppendChild(CreateWordItemStyles.CreateText(wordItem));
            }
            Console.WriteLine();
        }
        return body;
    }

    /// <summary>
    /// Removes bodies from Word document<br/>
    /// 0 = Removes all bodies<br/>
    /// Negative int = Removes n'th bodies in the back of document<br/>
    /// Positive int = Removes a specified body<br/>
    /// </summary>
    /// <param name="filepath"></param>
    /// <param name="remove"></param>
    public static void EditDoc(string filepath, int remove)
    {
        try
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true);
            if (wordDocument.MainDocumentPart != null && wordDocument.MainDocumentPart.Document.Body != null)
            {
                // Access the main document part
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                W.Body mainBody = mainPart.Document.Body;
                switch (remove)
                {
                    case 0:
                        mainBody.RemoveAllChildren();
                        break;
                    default:
                        if (remove < 0)
                        {
                            remove *= -1;
                            for (; remove > 0; remove--)
                                mainBody.LastChild?.Remove();
                        }
                        else
                            mainBody.ChildElements[remove].Remove();
                        break;
                }
                mainPart.Document.Save();
                Process.Start(new ProcessStartInfo(filepath) { UseShellExecute = true });
            }
        }
        catch
        {
            Console.WriteLine("Dokumentet er allerede åben. Luk word dokumentet!");
        }
    }

    /// <summary>
    /// Appends bodies to the Word Document<br/>
    /// </summary>
    /// <param name="filepath"></param>
    /// <param name="collection"></param>
    public static void EditDoc(string filepath, List<WordItem> collection)
    {
        try
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true);
            if (wordDocument.MainDocumentPart != null && wordDocument.MainDocumentPart.Document.Body != null)
            {
                // Access the main document part
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                W.Body mainBody = mainPart.Document.Body;
                mainBody.AppendChild(CreateBody(collection));
                Process.Start(new ProcessStartInfo(filepath) { UseShellExecute = true });
            }
        }
        catch
        {
            Console.WriteLine("Dokumentet er allerede åben. Luk word dokumentet! Edit");
        }
    }
}
class CreateWordItemStyles
{
    public static W.Paragraph CreateText(WordItem wordItem)
    {
        if (wordItem.MathField)
        {
            W.Paragraph mathParagraph;
            M.MathProperties mathProps = new();
            W.Run run = new(new W.Text(wordItem.Text));
            if (wordItem.XML)
                mathProps.AppendChild(new M.OfficeMath(run.InnerText));
            else
                mathProps.AppendChild(new M.OfficeMath(run));
            //Position
            M.ParagraphProperties ParaProp = new();
            switch (wordItem.Position)
            {
                case ParPos.Left:
                    ParaProp.Append(new M.Justification() { Val = M.JustificationValues.Left });
                    break;
                case ParPos.Center:
                    ParaProp.Append(new M.Justification() { Val = M.JustificationValues.Center });
                    break;
                case ParPos.Right:
                    ParaProp.Append(new M.Justification() { Val = M.JustificationValues.Right });
                    break;
            }

            mathParagraph = new(mathProps);
            mathParagraph.Append(ParaProp);
            return new(mathParagraph);
        }
        else
        {
            W.Run run = new(new W.Text(wordItem.Text));
            W.RunProperties runProps = new();
            //Position
            switch (wordItem.Position)
            {
                case ParPos.Left:
                    runProps.Append(new W.Justification() { Val = W.JustificationValues.Left });
                    break;
                case ParPos.Center:
                    runProps.Append(new W.Justification() { Val = W.JustificationValues.Center });
                    break;
                case ParPos.Right:
                    runProps.Append(new W.Justification() { Val = W.JustificationValues.Right });
                    break;
            }
            runProps.AppendChild(run);
            return new(runProps);
        }
    }
    public static W.Table CreateTable(List<WordItem> tableItems)
    {
        return new();
    }
}
class Program
{
    static void Main()
    {
        string fileName = "WordTest1.docx";
        string filePath = @"D:\CodeProjects\GitHub\WordPrint\WordDocuments";
        List<WordItem> wordItems = new();
        for (int i = 0; i < 10; i++)
        {
            wordItems.Add(new WordItem()
            {
                Text = i.ToString(),
                ItemType = WordItemType.Text,
                MathField = true,
            });
        }
        for (int i = 0; i < 10; i++)
        {
            wordItems.Add(new WordItem()
            {
                Text = i.ToString(),
                ItemType = WordItemType.Table,
            });
        }
        for (int i = 0; i < 10; i++)
        {
            wordItems.Add(new WordItem()
            {
                Text = i.ToString(),
                ItemType = WordItemType.Text,
                MathField = true,
            });
        }
        CreateDocument.EditDoc(Path.Combine(filePath, fileName), wordItems);
    }
}

/*
try
{
    using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true);
    if (wordDocument.MainDocumentPart != null && wordDocument.MainDocumentPart.Document.Body != null)
    {
        // Access the main document part
        MainDocumentPart mainPart = wordDocument.MainDocumentPart;
        W.Body MainBody = mainPart.Document.Body;
        switch (whichEdit)
        {
            case 0:
                //Add body to word document
                if (body != null)
                    MainBody.AppendChild(body);
                break;
            case 1:
                //Remove specific or last body from MainBody in word document
                if (RemoveThis != null)
                    MainBody.ChildElements[(int)RemoveThis].Remove();
                else
                    MainBody.LastChild?.Remove();
                break;
            case 2:
                //Clear all bodies from word document
                MainBody.RemoveAllChildren();
                break;
        }
        mainPart.Document.Save();
    }
    ProcessStartInfo psi = new(filepath)
    {
        UseShellExecute = true,
    };
    Process? process = Process.Start(psi);
}
catch
{
    Debug.WriteLine("Dokumentet er allerede åben. Luk word dokumentet!");
}
*/
