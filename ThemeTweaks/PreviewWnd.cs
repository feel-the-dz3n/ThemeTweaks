using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ThemeTweaks
{
    public partial class PreviewWnd : Form
    {
        Form1 f1 = null;
        public PreviewWnd(Form1 f)
        {
            f1 = f;
            InitializeComponent();
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void PreviewWnd_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            f1.UncheckPreview();
        }

        private void PreviewWnd_Load(object sender, EventArgs e)
        {
            Left = f1.Left + f1.Width;
            Top = f1.Top;

            if (!Aero.CheckAeroEnabled())
                button2.Enabled = false;

        }

        private void PreviewWnd_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible)
                PreviewWnd_Load(null, null);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Aero.Glass(this);
            this.AllowTransparency = true;
            this.TransparencyKey = Color.Black;
            this.BackColor = Color.Black;
        }
    }
}
