using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace InfoQuest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            string allArgs = "";
            for (int i = 0; i < args.Length; i++)
            {
                allArgs = allArgs + "Param:" + (i + 1) + " = " + args[i]+"\n";
            }
            MessageBox.Show(allArgs);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }
    }
}