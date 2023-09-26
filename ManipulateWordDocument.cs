using W = DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace WordPrint;

class WordItem
{
    public WordItem()
    {
        Content = "";
        Font = "";
        FontSize = 11;
        BackgroundColor = "";
        FontColor = "";
        Bold = false;
        Italic = false;
        Underline = TextUnderLine.None;
        ItemType = WordItemType.Text;
        MathField = false;
        XML = false;
        Position = ParPos.Left;
        Margin = new(0, 0, 0, 0);
        NewLine = true;
        Punctuation = 0;
        PunctuationTypesUnicode = new()
        {
            "\u2022",
            "\u2058",
            "\u2023",
            "\u2015",
            "\u2014",
            "\u2013",
            "\u2012",
            "\u2011",
            "\u2010",
        };
    }

    //Text Options
    public string Content;
    public string Font;
    public int FontSize;
    //Format Text
    public string BackgroundColor;
    public string FontColor;
    public bool Bold;
    public bool Italic;
    public TextUnderLine Underline;
    //Type
    public WordItemType ItemType;
    //Math
    public bool MathField;
    public bool XML;
    //Position
    public ParPos Position;
    public ParMargin Margin;
    public bool NewLine;
    //Punctuation
    public int Punctuation;
    public List<string> PunctuationTypesUnicode;
}
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
struct ParMargin
{
    public ParMargin(int left, int right, int top, int bottom)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
    }
    public ParMargin(int left, int right, int topBottom)
    {
        Left = left;
        Right = right;
        Top = topBottom;
        Bottom = topBottom;
    }
    public ParMargin(int top, int bottom)
    {
        Left = 0;
        Right = 0;
        Top = top;
        Bottom = bottom;
    }
    public ParMargin(int all)
    {
        Left = all;
        Right = all;
        Top = all;
        Bottom = all;
    }

    public int Left;
    public int Right;
    public int Top;
    public int Bottom;
}
struct CreateDocument
{
    /// <summary>
    /// Document Filepath<br/>
    /// </summary>
    public string? FilePath;

    /// <summary>
    /// Creates the document<br/>
    /// </summary>
    /// <param name="DocumentName"/>
    /// <param name="desktopPath"/>
    public static void CreateDoc(string DocumentName, string desktopPath)
    {
        System.Diagnostics.Stopwatch stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            if (!Directory.Exists(desktopPath))
                Directory.CreateDirectory(desktopPath);
            #region Create Word-Document
            using DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDoc = 
                DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(Path.Combine(desktopPath, $"{DocumentName}"), 
                DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            DocumentFormat.OpenXml.Packaging.MainDocumentPart MainPart = wordDoc.AddMainDocumentPart();
            MainPart.Document = new W.Document
            {
                Body = new W.Body()
            };
            MainPart.Document.Save();
            #endregion
        }
        catch (Exception e)
        {
            Errors.ErrorChecker(e);
        }
        stopwatch.Stop();
        CalculateTime.CreateCounterList.Add($"Created document in {stopwatch.ElapsedMilliseconds} ms");
        CalculateTime.CreateCount += (int)stopwatch.ElapsedMilliseconds;
    }

    /// <summary>
    /// Creates the body to append to main body<br/>
    /// </summary>
    /// <param name="collection"></param>
    public static W.Body CreateBody(List<WordItem> collection)
    {
        System.Diagnostics.Stopwatch stopwatch = System.Diagnostics.Stopwatch.StartNew();
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
        foreach ((WordItemType type, List<WordItem> items) in itemCollections)
        {
            if (type == WordItemType.Text)
            {
                List<List<WordItem>> lines = GetOneLines(items);
                foreach (List<WordItem> oneLine in lines)
                {
                    body.AppendChild(CreateWordItemStyles.CreateText(oneLine));
                }
            }
        }
        stopwatch.Stop();
        CalculateTime.CreateCounterList.Add($"Created body in {stopwatch.ElapsedMilliseconds} ms");
        CalculateTime.CreateCount += (int)stopwatch.ElapsedMilliseconds;
        return body;
    }

    /// <summary>
    /// Collect all lines and sort them in Run's.<br/>
    /// This can be changed with "NewLine".<br/>
    /// </summary>
    /// <param name="lines"/>
    private static List<List<WordItem>> GetOneLines(List<WordItem> lines)
    {
        List<List<WordItem>> collectionLines = new();
        List<WordItem> tempLine = new();
        foreach (WordItem line in lines)
        {
            if (line.NewLine && tempLine.Count > 0)
            {
                collectionLines.Add(tempLine);
                tempLine = new();
            }
            tempLine.Add(line);
        }
        collectionLines.Add(tempLine);
        return collectionLines;
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
        System.Diagnostics.Stopwatch stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            using DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filepath, true);
            if (wordDocument.MainDocumentPart != null && wordDocument.MainDocumentPart.Document.Body != null)
            {
                // Access the main document part
                DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart = wordDocument.MainDocumentPart;
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
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filepath) { UseShellExecute = true });
            }
        }
        catch (Exception e)
        {
            Errors.ErrorChecker(e);
        }
        stopwatch.Stop();
        CalculateTime.CreateCounterList.Add($"Edited document in {stopwatch.ElapsedMilliseconds} ms");
        CalculateTime.CreateCount += (int)stopwatch.ElapsedMilliseconds;
    }

    /// <summary>
    /// Appends bodies to the Word Document<br/>
    /// </summary>
    /// <param name="filepath"></param>
    /// <param name="collection"></param>
    public static void EditDoc(string filepath, List<WordItem> collection)
    {
        System.Diagnostics.Stopwatch stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            using DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filepath, true);
            if (wordDocument.MainDocumentPart != null && wordDocument.MainDocumentPart.Document.Body != null)
            {
                // Access the main document part
                DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                W.Body mainBody = mainPart.Document.Body;
                mainBody.AppendChild(CreateBody(collection));
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filepath) { UseShellExecute = true });
            }
        }
        catch (Exception e)
        {
            Errors.ErrorChecker(e);
        }
        stopwatch.Stop();
        CalculateTime.CreateCounterList.Add($"Edited document in {stopwatch.ElapsedMilliseconds} ms");
        CalculateTime.CreateCount += (int)stopwatch.ElapsedMilliseconds;
    }
}
struct CreateWordItemStyles
{
    public static List<(bool, List<WordItem>)> SplitMathAndText(List<WordItem> oneline)
    {
        List<(bool, List<WordItem>)> splitting = new();
        List<WordItem> tempSplit = new();
        foreach (var item in oneline)
        {
            if (tempSplit.Count > 0 && tempSplit[^1].MathField != item.MathField)
            {
                splitting.Add((tempSplit[0].MathField, tempSplit));
                tempSplit = new();
            }
            tempSplit.Add(item);
        }
        splitting.Add((tempSplit[0].MathField, tempSplit));
        return splitting;
    }

    public static W.Paragraph CreateText(List<WordItem> oneLine)
    {
        W.Paragraph paragraph = new();
        W.ParagraphProperties paraProp;
        List<(bool, List<WordItem>)> types = SplitMathAndText(oneLine);
        for (int i_1 = 0; i_1 < types.Count; i_1++)
        {
            if (types[i_1].Item1)
            {
                M.ParagraphProperties mathParaProp;
                M.Paragraph mathParagraph = new();
                for (int i_2 = 0; i_2 < types[i_1].Item2.Count; i_2++)
                {
                    #region M.Properties
                    mathParaProp = new()
                    {
                        Justification = new M.Justification()
                        {
                            Val = M.JustificationValues.Left,
                        }
                    };
                    mathParagraph.Append(mathParaProp);

                    #region Justification
                    M.Justification mathJustification = new();
                    switch (types[i_1].Item2[i_2].Position)
                    {
                        case ParPos.Left:
                            mathJustification.Val = M.JustificationValues.Left;
                            break;
                        case ParPos.Center:
                            mathJustification.Val = M.JustificationValues.Center;
                            break;
                        case ParPos.Right:
                            mathJustification.Val = M.JustificationValues.Right;
                            break;
                    }
                    mathParaProp.Append(mathJustification);
                    #endregion
                    #region Position
                    W.Indentation indentation = new()
                    {
                        LeftChars = types[i_1].Item2[i_2].Margin.Left,
                        RightChars = types[i_1].Item2[i_2].Margin.Right,
                    };
                    mathParaProp.Append(indentation);
                    #endregion
                    #region Punctuation
                    if (types[i_1].Item2[i_2].Punctuation > 0 && types[i_1].Item2[i_2].Punctuation < types[i_1].Item2[i_2].PunctuationTypesUnicode.Count)
                        types[i_1].Item2[i_2].Content = $"{types[i_1].Item2[i_2].PunctuationTypesUnicode[types[i_1].Item2[i_2].Punctuation - 1]}{types[i_1].Item2[i_2].Content}";
                    #endregion

                    #endregion

                    #region Paragraph
                    M.Run mathRun = new(new M.Text
                    {
                        Text = types[i_1].Item2[i_2].Content,
                    })
                    {
                        RunProperties = new W.RunProperties
                        {
                            RunFonts = new W.RunFonts
                            {
                                Ascii = types[i_1].Item2[i_2].Font,
                                HighAnsi = types[i_1].Item2[i_2].Font,
                            },
                            Italic = new W.Italic { Val = types[i_1].Item2[i_2].Italic },
                            Color = new W.Color { Val = types[i_1].Item2[i_2].FontColor },
                            Shading = new W.Shading { Fill = types[i_1].Item2[i_2].BackgroundColor }
                        },
                    };
                    if (types[i_1].Item2[i_2].XML)
                        mathParagraph.AppendChild(new M.OfficeMath(mathRun.InnerText));
                    else
                        mathParagraph.AppendChild(new M.OfficeMath(mathRun));
                    #endregion
                }
                paragraph.Append(mathParagraph);
            }
            else
            {
                for (int i_2 = 0; i_2 < types[i_1].Item2.Count; i_2++)
                {
                    #region W.Properties
                    paraProp = new();
                    paragraph.Append(paraProp);

                    #region Justification
                    W.Justification justification = new();
                    switch (types[i_1].Item2[i_2].Position)
                    {
                        case ParPos.Left:
                            justification.Val = W.JustificationValues.Left;
                            break;
                        case ParPos.Center:
                            justification.Val = W.JustificationValues.Center;
                            break;
                        case ParPos.Right:
                            justification.Val = W.JustificationValues.Right;
                            break;
                    }
                    paraProp.Append(justification);
                    #endregion
                    #region Position
                    W.Indentation indentation = new()
                    {
                        LeftChars = types[i_1].Item2[i_2].Margin.Left,
                        RightChars = types[i_1].Item2[i_2].Margin.Right,
                    };
                    paraProp.Append(indentation);
                    #endregion
                    #region Punctuation
                    if (types[i_1].Item2[i_2].Punctuation > 0 && types[i_1].Item2[i_2].Punctuation < types[i_1].Item2[i_2].PunctuationTypesUnicode.Count)
                        types[i_1].Item2[i_2].Content = $"{types[i_1].Item2[i_2].PunctuationTypesUnicode[types[i_1].Item2[i_2].Punctuation - 1]}   {types[i_1].Item2[i_2].Content}";
                    #endregion

                    #endregion

                    #region Paragraph
                    W.Run run = new(new W.Text
                    {
                        Text = types[i_1].Item2[i_2].Content,
                    })
                    {
                        RunProperties = new W.RunProperties
                        {
                            RunFonts = new W.RunFonts
                            {
                                Ascii = types[i_1].Item2[i_2].Font,
                                HighAnsi = types[i_1].Item2[i_2].Font,
                            },
                            Italic = new W.Italic { Val = types[i_1].Item2[i_2].Italic },
                            Color = new W.Color { Val = types[i_1].Item2[i_2].FontColor },
                            Shading = new W.Shading { Fill = types[i_1].Item2[i_2].BackgroundColor },
                        },
                    };
                    paragraph.Append(run);
                    #endregion
                }
            }
        }
        return paragraph;
    }
    public static W.Table CreateTable(List<WordItem> tableItems)
    {
        return new();
    }
}
struct Errors
{
    public static void ErrorChecker(Exception e)
    {
        switch (e.HResult)
        {
            case -2147024809:
                Console.WriteLine("Error code: -2147024809 \n" +
                                  "Error: Wrong unicode \n" +
                                 $"Value: {e.Message} \n" +
                                  "Fix: A WordItem's 'Content' property holds a value (Unicode), that cannot be used \n");
                break;
            case -2146233086:
                Console.WriteLine("Error code: -2146233086 \n" +
                                  "Error: Index was out of range \n" +
                                 $"Value: {e.Message} \n" +
                                  "Fix: Make sure no indexes can be out of range \n");
                break;
            default:
                Console.WriteLine("Document is open. \nClose the document, before it can be edited!");
                break;
        }
        Console.ReadKey();
    }
}
static class CalculateTime
{
    public static int CreateCount = 0;
    public static List<string> CreateCounterList = new();
    public static float TimeForLines(int lines)
    {
        return 0.03f * lines + 174.3f;
    }
}
