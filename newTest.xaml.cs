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
    public partial class newTest : Window
    {
        public newTest()
        {
            InitializeComponent();
        }
        private XpsDocument ConvertWordDocToXPSDocUpdated(word.Document document, string xpsDocName)
        {
            // You must ensure you have Microsoft.Office.Interop.Word.Dll version 12.
            // Version 11 or previous versions do not have WdSaveFormat.wdFormatXPS option
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

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {

            string docPath = GetPathFromCurrentProject(@"Rapports\Decision_PR.docx");
            string NewdocPath = GetPathFromCurrentProject(@"Rapports\Decision_PR (generated).docx");
            
            GenerateDocument(docPath,NewdocPath,(word.Application wordApp) =>
              {
                  FindAndReplace(wordApp, "<anne>", DateTime.Now.Year.ToString());
                  FindAndReplace(wordApp, "<societe>", "MonSociete");
                  FindAndReplace(wordApp, "<registreCommerce>", "MonRegistre");
                  FindAndReplace(wordApp, "<cnss>", "12345234");
                  FindAndReplace(wordApp, "<taxeProf>", "23423");
                  FindAndReplace(wordApp, "<numeroDemande>", "13423");
                  FindAndReplace(wordApp, "<domicile>", "324234");
                  FindAndReplace(wordApp, "<date>", $"{DateTime.Now.Day} / {DateTime.Now.Month} /{DateTime.Now.Year}");
              }
            ,documentViewer1);
        }
        private string GetPathFromCurrentProject(string FolderOrFileName)
        {
            return $@"{Directory.GetCurrentDirectory().Replace(@"\bin\Debug", "\\")}{FolderOrFileName}";
        }

        internal async void GenerateDocument(object filename, object SaveAsPath,Action<word.Application> FindAndReplace, DocumentViewer documentViewer = null)
        {
            Loading loading = new Loading();
            loading.Show();
            XpsDocument xpsDocument = await GenerateDocumentAsync(filename, SaveAsPath, FindAndReplace);
            if (documentViewer != null) documentViewer.Document = xpsDocument.GetFixedDocumentSequence();
            loading.Close();
        }
        //this method will find and replace the input words
        internal void FindAndReplace(word.Application wordApp, object ToFindText, object replaceWithText)
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


        //this method will create a new copy of the passed document word 
        /// <summary>
        /// this method will generate document and save it and convert it to XPS document and returnt in in a Task
        /// </summary>
        /// <param name="filename">the path of the document</param>
        /// <param name="SaveAsPath">the path where you want to save your document</param>
        /// <param name="FindAndReplace">this action allows to find and replace words in the document</param>
        internal Task<XpsDocument> GenerateDocumentAsync(object filename, object SaveAsPath, Action<word.Application> FindAndReplace)
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
                myWordDoc.SaveAs2(ref SaveAsPath, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                string newXPSDocumentName = String.Concat(Path.GetDirectoryName(filename.ToString()), "\\",
                                System.IO.Path.GetFileNameWithoutExtension(filename.ToString()), ".xps");

                XpsDocument xpsDocument = ConvertWordDocToXPSDocUpdated(myWordDoc, newXPSDocumentName);
                myWordDoc.Close();
                wordApp.Quit();
                return xpsDocument;
            });
        }
    }
}
