using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Xps.Packaging;
using GemBox.Document;
using System.IO;
using word=Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace Document_Viewr_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
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
        //async
        private Task<XpsDocument> ConvertWordDocToXPSDocAsync(string wordDocName, string xpsDocName)
        {
            return Task.Factory.StartNew(() =>
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
            });
        }
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            //Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            //// Set filter for file extension and default file extension
            //dlg.DefaultExt = ".docx";
            //dlg.Filter = "Word documents (.docx)|*.docx";
            //// Display OpenFileDialog by calling ShowDialog method
            //Nullable<bool> result = dlg.ShowDialog();
            //// Get the selected file name and display in a TextBox
            //if (result == true)
            //{
            //    if (dlg.FileName.Length > 0)
            //    {
            //        SelectedFileTextBox.Text = dlg.FileName;
            //        string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(dlg.FileName), "\\",
            //                       System.IO.Path.GetFileNameWithoutExtension(dlg.FileName), ".xps");
            //        // Set DocumentViewer.Document to XPS document
            //        documentViewer1.Document =
            //            ConvertWordDocToXPSDoc(dlg.FileName, newXPSDocumentName).GetFixedDocumentSequence();
            //    }
            //}

            string docPath = GetPathFromCurrentProject(@"Rapports\Decision_PR.docx");
            SelectedFileTextBox.Text = docPath;

            string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(docPath), "\\",

                            System.IO.Path.GetFileNameWithoutExtension(docPath), ".xps");
            // Set DocumentViewer.Document to XPS document
            //var task = new Task(() =>
            //{
            //       SetDocumentViewer(docPath, newXPSDocumentName);
            //});
            //task.Start();
            SetDocumentViewer(docPath,newXPSDocumentName);
        }
        private  void SetDocumentViewer(string docPath,string newXPSDocumentName)
        {
            //Loading loading = new Loading();
            //loading.Show();
            XpsDocument document = ConvertWordDocToXPSDoc(docPath, newXPSDocumentName);
            documentViewer1.Document = document.GetFixedDocumentSequence();
            //loading.Close();
    //ConvertWordDocToXPSDoc(docPath, newXPSDocumentName).GetFixedDocumentSequence();
        }
        private string GetPathFromCurrentProject(string FolderOrFileName)
        {
            return $@"{Directory.GetCurrentDirectory().Replace(@"\bin\Debug", "\\")}{FolderOrFileName}";
        }
    }
}
