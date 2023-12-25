// Microsoft.Office.Interop.Word
// dotnet add package Microsoft.Office.Interop.Word --version 15.0.4797.1004

using System;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    static void Main()
    {
        string docPath = @"test_in.docx"; // Word文書のパス
        string pattern = @"abcd"; // 検索する正規表現パターン

        var wordApp = new Word.Application();
        try
        {
            Word.Document doc = wordApp.Documents.Open(docPath);
            Word.Range docRange = doc.Content;

            Word.Find find = docRange.Find;
            find.Text = pattern;
            find.MatchWildcards = true; // 正規表現を使う場合はこれをtrueに設定

            while (find.Execute())
            {
                int pageNumber = (int)docRange.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                int lineNum = (int)docRange.get_Information(Word.WdInformation.wdFirstCharacterLineNumber);
                Console.WriteLine($"Match found on Page {pageNumber}, Line {lineNum}");
                
                // 次の検索対象範囲を設定
                int rangeEnd = docRange.End;
                docRange = doc.Range(rangeEnd, doc.Content.End);
            }

            doc.Close();
        }
        finally
        {
            wordApp.Quit();
        }
    }
}

