using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    class Utility
    {
        public static StringBuilder Concat(StringBuilder sb, params object[] args)
        {
            sb.Append(string.Concat(args));
            return sb.Append('\n');
        }

        public static string Concat(params object[] args)
        {
            return string.Concat(string.Concat(args), '\n');
        }

        public static void ShowException(Exception ex)
        {
            string msg = string.Concat(ex.Message, "\n", ex.StackTrace);
            MessageBox.Show(msg);
        }

        public static bool CheckPPTFile(string fileName, StringBuilder sb)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                Utility.Concat(sb, "Please select a PowerPoint file.");
                return false;
            }
            else if (!File.Exists(fileName))
            {
                Utility.Concat(sb, "Could not open the following file: \n \"", fileName, "\"\nMake sure the file path you specified above is correct.");
                return false;
            }
            return true;
        }

        public static void StarProcess(string silverLightHtmlFile)
        {
            Process p = new Process();
            p.StartInfo.FileName = @"C:\Program Files\Internet Explorer\iexplore.exe";
            p.StartInfo.Arguments = silverLightHtmlFile;

            p.Start();
        }
    }
}
