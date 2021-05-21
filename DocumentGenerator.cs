using System;
using System.Windows;
using System.Windows.Xps.Packaging;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Document_Viewr_WPF
{
    class DocumentGenerator
    {
        internal static XpsDocument ConvertWordDocToXPSDocUpdated(word.Document document, string xpsDocName)
        {
            try
            {
                document.SaveAs(xpsDocName, word.WdSaveFormat.wdFormatXPS);
                XpsDocument xpsDoc = new XpsDocument(xpsDocName, System.IO.FileAccess.Read);
                return xpsDoc;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
            return null;
        }
        internal static async void GenerateDocument(object filename, Action<word.Application> FindAndReplace, DocumentViewer documentViewer = null)
        {
            Loading loading = new Loading();
            loading.Show();
            XpsDocument xpsDocument = await GetXpsDocumentAsync(filename, FindAndReplace);
            if (documentViewer != null) documentViewer.Document = xpsDocument.GetFixedDocumentSequence();
            //I wont need this for know cause by default when i want the close document word they are asking to save As
            //await CreateNewDocumentAsync(filename, FindAndReplace);
            loading.Close();
        }
        internal static void FindAndReplace(word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        internal static Task<XpsDocument> GetXpsDocumentAsync(object filename,Action<word.Application> FindAndReplace)
        {
            return Task.Factory.StartNew(() =>
            {
                word.Application wordApp = new word.Application();
                object missing = Missing.Value;
                word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;
                    wordApp.Documents.Add(filename);
                    myWordDoc = wordApp.ActiveDocument;

                    FindAndReplace(wordApp);
                }
                else
                {
                    return null;
                }
                string newXPSDocumentName = String.Concat(Path.GetDirectoryName(filename.ToString()), "\\",
                                System.IO.Path.GetFileNameWithoutExtension(filename.ToString()), ".xps");
                XpsDocument xpsDocument=null;
                try
                {
                     xpsDocument= ConvertWordDocToXPSDocUpdated(myWordDoc, newXPSDocumentName);
                    myWordDoc.Close();
                }catch(Exception exp)
                {
                    MessageBox.Show("La document n'est pas sauvegarder");
                }
                wordApp.Quit();
                return xpsDocument;
            });
        }
        internal static Task CreateNewDocumentAsync(object filename, Action<word.Application> FindAndReplace)
        {
            return Task.Factory.StartNew(() =>
            {
                word.Application wordApp = new word.Application();
                object missing = Missing.Value;
                word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;
                    wordApp.Documents.Add(filename);
                    myWordDoc = wordApp.ActiveDocument;

                    FindAndReplace(wordApp);
                }
                else
                {
                    return;
                }
                myWordDoc.SaveAs2(ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);
                myWordDoc.Close();
                wordApp.Quit();
            });
        }
    }
}
