using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Document_Viewr_WPF
{
    class RapportPath
    {
        public string Value { get; set; }
        public RapportPath(string value)
        {
            this.Value = value;
        }
        public static RapportPath Decision_PR { get { return new RapportPath(GetPathFromCurrentProject("Rapports\\Decision_PR.docx")); } }

        private static string GetPathFromCurrentProject(string FolderOrFileName)
        {
            return $@"{Directory.GetCurrentDirectory().Replace(@"\bin\Debug", "\\")}{FolderOrFileName}";
        }

    }
}
