using Microsoft.Office.Core;
using Microsoft.Office.Interop.Access.Dao;
using System.Diagnostics.Metrics;
using System.Reflection.Metadata;
using Word = Microsoft.Office.Interop.Word;

// Open word app and make invisible (leave on if you want to see file open)
var wordApp = new Word.Application();
wordApp.Visible = false;

// So the working directory for this is bin (there's a test file in bin) - change below as needed
// Opens the word document for testing
var docx = wordApp.Documents.Open($"{Directory.GetCurrentDirectory()}\\test.docx");

/*
 * I'm not sure we want to know the position (you can use that and string len to delete then replace)
 * But there are other ways
 * 
 * Currently this just finds a red (classic red or dark red) colored text and then displays it.
 */

int cnt = 0;
foreach (Word.Range r in docx.Content.Words)
{
    cnt++;
    if (r.Font.ColorIndex == Word.WdColorIndex.wdRed ||
        r.Font.ColorIndex == Word.WdColorIndex.wdClassicRed ||
        r.Font.ColorIndex == Word.WdColorIndex.wdDarkRed)
    {
        Console.WriteLine("\n---------------------");
        Console.WriteLine($"Red word at {cnt}");
        Console.WriteLine($"Replace {r.Text}");
        Console.WriteLine("---------------------\n");
    }
    else
    {
        Console.WriteLine(r.Text);
    }
}

// Close the document
docx.Close();
