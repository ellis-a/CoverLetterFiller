using Microsoft.Office.Interop.Word;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

internal class Program
{
    private static void Main(string[] args)
    {
        try
        {
            var loc = Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString();
            var replaceFields = File.ReadAllLines(Path.Combine(loc, "ReplaceFields.txt"));

            Application app = new();
            Document doc = app.Documents.Open(Path.Combine(loc, "Template.docx"));
            var newFileName = Path.Combine(loc, $"Cover Letter {DateTime.Now:yyyy-MMM-dd HH-mm-ss}.docx");

            doc.SaveAs2(newFileName);
            doc.Close();

            var fieldDict = new Dictionary<string, string>();

            foreach (var field in replaceFields)
            {
                Console.WriteLine($"{field}:");
                var replacement = Console.ReadLine();
                fieldDict.Add($"[{field}]", replacement ?? string.Empty);
            }

            Application wordApp = new() { Visible = false };
            Document aDoc = wordApp.Documents.Open(newFileName, ReadOnly: false, Visible: false);
            aDoc.Activate();

            foreach (var kvp in fieldDict)
            {
                FindAndReplace(wordApp, kvp.Key, kvp.Value);
            }

            aDoc.Save();
            aDoc.Close();

            Console.WriteLine("Done!");
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            Console.ReadLine(); 
        }
    }

    private static void ReplaceHeader(Application wordApp, Document doc)
    {
        // Loop through all sections
        foreach (Section section in doc.Sections)
        {
            doc.TrackRevisions = false; //Disable Tracking for the Field replacement operation

            //Get all Headers
            HeadersFooters headers = section.Headers;

            //Section headerfooter loop for all types enum WdHeaderFooterIndex. wdHeaderFooterEvenPages/wdHeaderFooterFirstPage/wdHeaderFooterPrimary;                          
            foreach (HeaderFooter header in headers)
            {
                Fields fields = header.Range.Fields;

                foreach (Field field in fields)
                {
                    field.Select();
                    field.Delete();
                    wordApp.Selection.TypeText("[DATE]");
                }
            }
        }
    }

    private static void FindAndReplace(Application doc, object findText, object replaceWithText)
    {
        //options
        object matchCase = true;
        object matchWholeWord = true;
        object matchWildCards = false;
        object matchSoundsLike = false;
        object matchAllWordForms = false;
        object forward = true;
        object format = false;
        object matchKashida = false;
        object matchDiacritics = false;
        object matchAlefHamza = false;
        object matchControl = false;
        object read_only = false;
        object visible = true;
        object replace = 2;
        object wrap = 1;
        //execute find and replace
        doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
            ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
    }
}