using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ThemeTweaks
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
            if(MessageBox.Show("AAAAAAA.\r\n\r\nPlease, do not use this program if you do not know what it does.\r\nI (author) am not responsible for the consequences!!1111\r\n\r\nDo you want to open this program?", "ThemeTweaks disclaimer", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                Environment.Exit(0);
            }
            Application.Run(new Form1());
        }
    }
}
