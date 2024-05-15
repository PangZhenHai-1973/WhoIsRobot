using System;
using System.IO;
using System.Windows.Forms;

namespace WhoIsRobot
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            FileInfo FI = new FileInfo(Application.ExecutablePath);
            string s = FI.Name;
            s = s.Substring(0, s.Length - 4);
            System.Diagnostics.Process[] processes1 = System.Diagnostics.Process.GetProcessesByName(s);
            if (processes1.Length > 1)
            {
                Application.Exit();
            }
            else
            {
                ApplicationConfiguration.Initialize();
                Application.Run(new Form1());
            }
        }
    }
}