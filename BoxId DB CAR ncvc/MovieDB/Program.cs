using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace BoxIdDb
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmLogin(Assembly.GetExecutingAssembly().GetName().Name));
        }
    }
}