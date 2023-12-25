// Microsoft.Office.Interop.Word
// dotnet add package Microsoft.Office.Interop.Word --version 15.0.4797.1004

using System;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    static void Main()
    {
        // Word文書のパス
        string docPath = @".\test_in.docx";

        // Wordアプリケーションを開始
        var wordApp = new Word.Application();
        wordApp.Visible = false;

        // 文書を開く
        Word.Document doc = wordApp.Documents.Open(docPath);

        try
        {
            Word.Range range = doc.Content;
            int pageCount = wordApp.ActiveWindow.ActivePane.Pages.Count;

            for (int i = 1; i <= pageCount; i++)
            {
                range.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, i);
                int pageNumber = (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                Console.WriteLine($"Page Number: {pageNumber}");
            }
        }
        finally
        {
            // 文書を閉じる
            doc.Close();
            wordApp.Quit();
        }
    }
}
