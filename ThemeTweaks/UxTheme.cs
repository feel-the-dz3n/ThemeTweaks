using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ThemeTweaks
{
    public static class UxTheme
    {
        [DllImport("UxTheme.Dll", EntryPoint = "#65", CharSet = CharSet.Unicode)]
        public static extern int SetSystemVisualStyle(string pszFilename, string pszColor, string pszSize, int dwReserved);

        [DllImport("UxTheme.Dll")]
        public static extern int EnableTheming(bool fEnable);
    }
}
