using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    static class MainClass
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new PowerPointReaderForm());
        }
    }
}