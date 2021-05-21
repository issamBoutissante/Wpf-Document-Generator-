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
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {

            DocumentGenerator.GenerateDocument(RapportPath.Decision_PR.Value,(word.Application wordApp) =>
              {
                  DocumentGenerator.FindAndReplace(wordApp, "<anne>", DateTime.Now.Year.ToString());
                  DocumentGenerator.FindAndReplace(wordApp, "<societe>", "MonSociete");
                  DocumentGenerator.FindAndReplace(wordApp, "<registreCommerce>", "MonRegistre");
                  DocumentGenerator.FindAndReplace(wordApp, "<cnss>", "12345234");
                  DocumentGenerator.FindAndReplace(wordApp, "<taxeProf>", "23423");
                  DocumentGenerator.FindAndReplace(wordApp, "<numeroDemande>", "13423");
                  DocumentGenerator.FindAndReplace(wordApp, "<domicile>", "324234");
                  DocumentGenerator.FindAndReplace(wordApp, "<date>", $"{DateTime.Now.Day} / {DateTime.Now.Month} /{DateTime.Now.Year}");
              }
            ,documentViewer1);
        }
    }
}
