using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Xps.Packaging;
using GemBox.Document;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.Reflection;

namespace Document_Viewr_WPF
{
    public partial class newTest : Window
    {
        public newTest()
        {
            InitializeComponent();
        }
        #region ConvertWordDocToXPSDoc Sync
        private XpsDocument ConvertWordDocToXPSDoc(string wordDocName, string xpsDocName)
        {
            // Create a WordApplication and add Document to it
            word.Application
                wordApplication = new word.Application();

            wordApplication.Documents.Add(wordDocName);
            word.Document doc = wordApplication.ActiveDocument;
            // You must ensure you have Microsoft.Office.Interop.Word.Dll version 12.
            // Version 11 or previous versions do not have WdSaveFormat.wdFormatXPS option
            try
            {
                doc.SaveAs(xpsDocName, word.WdSaveFormat.wdFormatXPS);
                wordApplication.Quit();
                XpsDocument xpsDoc = new XpsDocument(xpsDocName, System.IO.FileAccess.Read);
                return xpsDoc;
            }
            catch (Exception exp)
            {
                string str = exp.Message;
            }
            return null;
        }
        private XpsDocument ConvertWordDocToXPSDocUpdated(word.Document document, string xpsDocName)
        {
            // Create a WordApplication and add Document to it

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
                string str = exp.Message;
            }
            return null;
        }
        #endregion
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            string docPath = GetPathFromCurrentProject(@"Rapports\Decision_PR.docx");
            string NewdocPath = GetPathFromCurrentProject(@"Rapports\Decision_PR (generated).docx");
            CreateWordDocument(docPath,NewdocPath,(word.Application wordApp) =>
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
              );
        }
        //private async void SetDocumentViewer(string docPath, string newXPSDocumentName)
        //{
        //    Loading loading = new Loading();
        //    loading.Show();
        //    XpsDocument document = await ConvertWordDocToXPSDocAsync(docPath, newXPSDocumentName);
        //    documentViewer1.Document = document.GetFixedDocumentSequence();
        //    loading.Close();
        //    //ConvertWordDocToXPSDoc(docPath, newXPSDocumentName).GetFixedDocumentSequence();
        //}
        private string GetPathFromCurrentProject(string FolderOrFileName)
        {
            return $@"{Directory.GetCurrentDirectory().Replace(@"\bin\Debug", "\\")}{FolderOrFileName}";
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
        internal void CreateWordDocument(object filename, object SaveAsPath, Action<word.Application> FindAndReplace)
        {
            MessageBox.Show("Start");
            word.Application wordApp = new word.Application();
            MessageBox.Show("Application a ete cree");
            object missing = Missing.Value;
            word.Document myWordDoc = null;
            MessageBox.Show("Document a ete cree");

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;
                MessageBox.Show("start open");
                wordApp.Documents.Add(filename);
                myWordDoc = wordApp.ActiveDocument;

                MessageBox.Show("open done");
                //myWordDoc.Activate();
                //find and replace
                MessageBox.Show("start Find and replace");
                FindAndReplace(wordApp);
                MessageBox.Show("start Find and replace done");
                //inserted
                //UpdatedemyWordDoc             //Save the document
                MessageBox.Show("Set Document a ete effectue");
            }
            else
            {
                MessageBox.Show("Document pas trouve!");
                return;
            }
            myWordDoc.SaveAs2(ref SaveAsPath, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            //object docPath = GetPathFromCurrentProject(@"Rapports\Decision_PR.docx");

            string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(SaveAsPath.ToString()), "\\",

                            System.IO.Path.GetFileNameWithoutExtension(SaveAsPath.ToString()), ".xps");


            word.Application wordApplication = new word.Application();

            wordApplication.Documents.Add(SaveAsPath);
            word.Document doc = wordApplication.ActiveDocument;

            XpsDocument document = ConvertWordDocToXPSDoc(SaveAsPath.ToString(), newXPSDocumentName);
            documentViewer1.Document = document.GetFixedDocumentSequence();
            doc.Close();
            wordApplication.Quit();
            MessageBox.Show("Document A ete cree!");
        }
    }
}
