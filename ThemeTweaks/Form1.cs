using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ThemeTweaks
{
    public partial class Form1 : Form
    {
        public ServiceController ThemesService = null;
        public ServiceController DWMService = null;
        public PreviewWnd pWnd = null;
        public List<Control> ThemesEdited = new List<Control>();
        public List<Control> DWMEdited = new List<Control>();
        public List<Control> CTEdited = new List<Control>();
        public bool IsUpdating = false;
        public string CurrentThemePath { get { return tbCurrentTheme.Text; } set { tbCurrentTheme.Text = value; } }

        public RegistryKey SarahWhereIsMyTea
        {
            get
            {
                if (cbDefUser.Checked)
                    return Registry.Users.OpenSubKey(".DEFAULT");
                else if (cbHKCU.Checked)
                    return Registry.CurrentUser;
                else if (cbHKLM.Checked)
                    return Registry.LocalMachine;
                else
                    return null;
            }
        }
        public Form1()
        {
            InitializeComponent();
            try
            {
                ThemesService = new ServiceController("themes");
            }
            catch (Exception ex)
            {
                gbThemesService.Enabled = false;
                MessageBox.Show(ex.ToString(), "Can't access Themes service");
            }

            try
            {
                DWMService = new ServiceController("uxsms");
            }
            catch (Exception ex)
            {
                gbDWMService.Enabled = false;
                MessageBox.Show(ex.ToString(), "Can't access UXSMS service");
            }
        }

        public void InvokedStatus()
        {
            labelEditedDWM.Text = "Edited elements: " + DWMEdited.Count;
            labelEditedThemes.Text = "Edited elements: " + ThemesEdited.Count;
            labelEditedCurrentTheme.Text = "Edited elements: " + CTEdited.Count;
            if (ThemesService != null)
            {
                ThemesService.Refresh();

                labelThemesStatus.Text = ThemesService.Status.ToString();
                if (ThemesService.Status == ServiceControllerStatus.Stopped)
                    btnThemesStopStart.Text = "Start";
                else if (ThemesService.Status == ServiceControllerStatus.Running)
                    btnThemesStopStart.Text = "Stop";
                else
                    btnThemesStopStart.Text = "...";
            }

            if (DWMService != null)
            {
                DWMService.Refresh();

                labelUxsmsStatus.Text = DWMService.Status.ToString();
                if (DWMService.Status == ServiceControllerStatus.Stopped)
                    btnUxsmsStopStart.Text = "Start";
                else if (DWMService.Status == ServiceControllerStatus.Running)
                    btnUxsmsStopStart.Text = "Stop";
                else
                    btnUxsmsStopStart.Text = "...";
            }

            IsUpdating = true;
            UpdateReg();
            UpdateCT();
            IsUpdating = false;
        }

        public void GetStatus()
        {
            while (true)
            {

                Invoke(new Action(() =>
                {
                    InvokedStatus();
                }));
                Thread.Sleep(500);
            }
        }

        public void UncheckPreview()
        {
            checkBox1.Checked = false;
        }

        private void cbThemes_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        public static string GetDefaultThemePath()
        {
            if (Environment.OSVersion.Version.Major == 5) // Windows XP
                return @"%SystemRoot%\resources\Themes\Luna\Luna.msstyles";
            else // Other Versions
                return @"%SystemRoot%\resources\Themes\Aero\Aero.msstyles";
        }

        public void UpdateCT()
        {
            if (tabControl1.SelectedTab != tabPageCurrentTheme)
                return;
            try
            {
                if (CurrentThemePath.Length >= 1)
                {
                    gbCurrentTheme.Text = CurrentThemePath;
                    gbCurrentTheme.Enabled = true;

                    string[] c = { };
                    using (StreamReader reader = new StreamReader(CurrentThemePath))
                    {
                        c = reader.ReadToEnd().Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    }

                    for (int i = 0; i < c.Length; i++)
                    {
                        if (c[i].StartsWith("[VisualStyles]"))
                            for (int x = i; x < c.Length; x++)
                            {
                                i = x;
                                if (c[i].Length <= 1)
                                {
                                    break;
                                }
                                if (c[i].StartsWith("Path=") && !CTEdited.Contains(tbPathCurrentTheme))
                                    tbPathCurrentTheme.Text = c[i].Split('=')[1];

                                if (c[i].StartsWith("ColorStyle=") && !CTEdited.Contains(tbColorStyleCurrentTheme))
                                    tbColorStyleCurrentTheme.Text = c[i].Split('=')[1];

                                if (c[i].StartsWith("VisualStyleVersion=") && !CTEdited.Contains(tbVisualStyleVersionCurrentTh))
                                    tbVisualStyleVersionCurrentTh.Text = c[i].Split('=')[1];

                                if (c[i].StartsWith("Transparency=") && !CTEdited.Contains(cbTransparencyCurrentTheme))
                                {
                                    if (c[i].Split('=')[1] == "1")
                                    {
                                        cbTransparencyCurrentTheme.CheckState = CheckState.Checked;
                                        cbTransparencyCurrentTheme.Checked = true;
                                    }
                                    else if (c[i].Split('=')[1] == "0")
                                    {
                                        cbTransparencyCurrentTheme.CheckState = CheckState.Unchecked;
                                        cbTransparencyCurrentTheme.Checked = false;
                                    }
                                }

                                if (c[i].StartsWith("Composition=") && !CTEdited.Contains(cbComposCurrentTheme))
                                {
                                    if (c[i].Split('=')[1] == "1")
                                    {
                                        cbComposCurrentTheme.CheckState = CheckState.Checked;
                                        cbComposCurrentTheme.Checked = true;
                                    }
                                    else if (c[i].Split('=')[1] == "0")
                                    {
                                        cbComposCurrentTheme.CheckState = CheckState.Unchecked;
                                        cbComposCurrentTheme.Checked = false;
                                    }
                                }
                            }
                    }
                }
                else
                {
                    gbCurrentTheme.Text = "No Theme Wtf";
                    gbCurrentTheme.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                gbCurrentTheme.Text = ex.ToString();
                gbCurrentTheme.Enabled = false;
            }
        }

        public void UpdateReg()
        {
            using (RegistryKey ThemeManager = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager"))
            {
                if(ThemeManager == null)
                {
                    groupBoxThemeManager.Enabled = false;
                    goto UpdateReg2;
                }
                groupBoxThemeManager.Text = ThemeManager.Name;

                try
                {
                    if (!ThemesEdited.Contains(tbDllName))
                        tbDllName.Text = ThemeManager.GetValue("DllName").ToString();
                }
                catch
                {
                    tbDllName.Text = "";
                }

                try
                {
                    if (!ThemesEdited.Contains(cbLoadedBefore))
                    {
                        if (ThemeManager.GetValue("LoadedBefore").ToString() == "0")
                            cbLoadedBefore.Checked = false;
                        else if (ThemeManager.GetValue("LoadedBefore").ToString() == "1")
                            cbLoadedBefore.Checked = true;
                        else
                            throw new Exception();
                    }
                }
                catch
                {
                    cbLoadedBefore.CheckState = CheckState.Indeterminate;
                }

                try
                {
                    if (!ThemesEdited.Contains(cbThemeActive))
                    {
                        if (ThemeManager.GetValue("ThemeActive").ToString() == "0")
                            cbThemeActive.Checked = false;
                        else if (ThemeManager.GetValue("ThemeActive").ToString() == "1")
                            cbThemeActive.Checked = true;
                        else
                            throw new Exception();
                    }
                }
                catch
                {
                    cbThemeActive.CheckState = CheckState.Indeterminate;
                }

            }
            UpdateReg2:
            if (cbDefUser.Checked)
                gbThemes.Enabled = false;
            else
            {
                gbThemes.Enabled = true;
                using (RegistryKey Themes = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Themes"))
                {
                    if (Themes == null)
                    {
                        gbThemes.Enabled = false;
                        goto UpdateReg3;
                    }
                    gbThemes.Text = Themes.Name;


                    try
                    {
                        if (!ThemesEdited.Contains(tbCurrentTheme))
                            tbCurrentTheme.Text = Themes.GetValue("CurrentTheme").ToString();
                    }
                    catch
                    {
                        tbCurrentTheme.Text = "";
                    }

                    try
                    {
                        if (!ThemesEdited.Contains(tbInstallVisualStyle))
                            tbInstallVisualStyle.Text = Themes.GetValue("InstallVisualStyle").ToString();
                    }
                    catch
                    {
                        tbInstallVisualStyle.Text = "";
                    }


                    try
                    {
                        if (!ThemesEdited.Contains(cbDropShadow))
                        {
                            if (Themes.GetValue("Drop Shadow").ToString() == "FALSE")
                                cbDropShadow.Checked = false;
                            else if (Themes.GetValue("Drop Shadow").ToString() == "TRUE")
                                cbDropShadow.Checked = true;
                            else
                                throw new Exception();
                        }
                    }
                    catch
                    {
                        cbDropShadow.CheckState = CheckState.Indeterminate;
                    }


                    try
                    {
                        if (!ThemesEdited.Contains(cbFlatMenus))
                        {
                            if (Themes.GetValue("Flat Menus").ToString() == "FALSE")
                                cbFlatMenus.Checked = false;
                            else if (Themes.GetValue("Flat Menus").ToString() == "TRUE")
                                cbFlatMenus.Checked = true;
                            else
                                throw new Exception();
                        }


                    }
                    catch
                    {
                        cbFlatMenus.CheckState = CheckState.Indeterminate;
                    }

                }
            }
            UpdateReg3:
            if (cbHKLM.Checked)
            {
                gbDWM.Enabled = false;
            }
            else
            {
                gbDWM.Enabled = true;
                using (RegistryKey DWM = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\DWM"))
                {
                    if (DWM == null)
                    {
                        gbDWM.Enabled = false;
                        goto UpdateReg3;
                    }
                    gbDWM.Text = DWM.Name;

                    try
                    {
                        if (!DWMEdited.Contains(cbDWMCompos))
                        {
                            if (DWM.GetValue("Composition").ToString() == "0")
                                cbDWMCompos.Checked = false;
                            else if (DWM.GetValue("Composition").ToString() == "1")
                                cbDWMCompos.Checked = true;
                            else
                                throw new Exception();
                        }
                    }
                    catch
                    {
                        cbDWMCompos.CheckState = CheckState.Indeterminate;
                    }



                    try
                    {
                        if (!DWMEdited.Contains(tbDWMColorizationColor))
                            tbDWMColorizationColor.Text = ((int)DWM.GetValue("ColorizationColor")).ToString("X4");
                    }
                    catch
                    {
                        tbDWMColorizationColor.Text = "";
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            new Thread(GetStatus).Start();

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        private void btnThemesStopStart_Click(object sender, EventArgs e)
        {
            ServiceController service = ThemesService;
            try
            {
                if (service.Status == ServiceControllerStatus.Stopped)
                    service.Start();
                else if (service.Status == ServiceControllerStatus.Running)
                    service.Stop();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can't change status");
            }
            InvokedStatus();
        }

        private void btnThemesRestart_Click(object sender, EventArgs e)
        {
            ServiceController service = ThemesService;
            try
            {
                if (service.Status == ServiceControllerStatus.Running)
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped);
                }

                if (service.Status == ServiceControllerStatus.Stopped)
                    service.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can't change status");
            }
            InvokedStatus();
        }

        private void btnUxsmsStopStart_Click(object sender, EventArgs e)
        {
            ServiceController service = DWMService;
            try
            {
                if (service.Status == ServiceControllerStatus.Stopped)
                    service.Start();
                else if (service.Status == ServiceControllerStatus.Running)
                    service.Stop();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can't change status");
            }
            InvokedStatus();
        }

        private void labelUxSmsRestart_Click(object sender, EventArgs e)
        {
            ServiceController service = DWMService;
            try
            {
                if (service.Status == ServiceControllerStatus.Running)
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped);
                }

                if (service.Status == ServiceControllerStatus.Stopped)
                    service.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can't change status");
            }
            InvokedStatus();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (pWnd == null)
                pWnd = new PreviewWnd(this);

            if (checkBox1.Checked)
                pWnd.Show();
            else
            {
                pWnd.Dispose();
                pWnd = null;
            }
        }

        private void btnThemesRefresh_Click(object sender, EventArgs e)
        {
            ThemesEdited.Clear();
        }

        private void btnThemesApply_Click(object sender, EventArgs e)
        {
            try
            {
                using (RegistryKey ThemeManager = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager", true))
                {
                    ThemeManager.SetValue("DllName", tbDllName.Text, RegistryValueKind.ExpandString);

                    if (cbLoadedBefore.CheckState != CheckState.Indeterminate)
                    {
                        if (cbLoadedBefore.Checked)
                            ThemeManager.SetValue("LoadedBefore", 1, RegistryValueKind.String);
                        else if (!cbLoadedBefore.Checked)
                            ThemeManager.SetValue("LoadedBefore", 0, RegistryValueKind.String);
                    }

                    if (cbThemeActive.CheckState != CheckState.Indeterminate)
                    {
                        if (cbThemeActive.Checked)
                            ThemeManager.SetValue("ThemeActive", 1, RegistryValueKind.String);
                        else if (!cbThemeActive.Checked)
                            ThemeManager.SetValue("ThemeActive", 0, RegistryValueKind.String);
                    }
                }
                using (RegistryKey Themes = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Themes", true))
                {
                    Themes.SetValue("CurrentTheme", tbCurrentTheme.Text, RegistryValueKind.String);
                    Themes.SetValue("InstallVisualStyle", tbInstallVisualStyle.Text, RegistryValueKind.ExpandString);

                    if (cbDropShadow.CheckState != CheckState.Indeterminate)
                    {
                        if (cbDropShadow.Checked)
                            Themes.SetValue("Drop Shadow", "TRUE", RegistryValueKind.String);
                        else if (!cbDropShadow.Checked)
                            Themes.SetValue("Drop Shadow", "FALSE", RegistryValueKind.String);
                    }

                    if (cbFlatMenus.CheckState != CheckState.Indeterminate)
                    {
                        if (cbFlatMenus.Checked)
                            Themes.SetValue("Flat Menus", "TRUE", RegistryValueKind.String);
                        else if (!cbFlatMenus.Checked)
                            Themes.SetValue("Flat Menus", "FALSE", RegistryValueKind.String);
                    }
                }

                // ...
                ThemesEdited.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (cbAutoRestart.Checked)
                btnThemesRestart_Click(null, null);
        }

        private void btnDWMRefresh_Click(object sender, EventArgs e)
        {
            DWMEdited.Clear();
        }

        private void btmDWMApply_Click(object sender, EventArgs e)
        {
            try
            {
                using (RegistryKey DWM = SarahWhereIsMyTea.OpenSubKey("Software\\Microsoft\\Windows\\DWM", true))
                {
                    if (cbDWMCompos.Checked)
                        DWM.SetValue("Composition", 1, RegistryValueKind.DWord);
                    else if (!cbDWMCompos.Checked)
                        DWM.SetValue("Composition", 0, RegistryValueKind.DWord);

                    if (tbDWMColorizationColor.TextLength >= 1)
                        DWM.SetValue("ColorizationColor", int.Parse(tbDWMColorizationColor.Text, System.Globalization.NumberStyles.HexNumber), RegistryValueKind.DWord);
                }
                // ...
                DWMEdited.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (cbDWMRestart.Checked)
                labelUxSmsRestart_Click(null, null);
        }

        private void btnStopBoth_Click(object sender, EventArgs e)
        {
            try
            {
                DWMService.Stop();
            }
            catch { }
            try
            {
                ThemesService.Stop();
            }
            catch { }
        }

        private void btnStartBoth_Click(object sender, EventArgs e)
        {
            try
            {
                DWMService.Start();
            }
            catch { }
            try
            {
                ThemesService.Start();
            }
            catch { }
        }

        private void btnRestartBoth_Click(object sender, EventArgs e)
        {
            try
            {
                ThemesService.Stop();
                ThemesService.WaitForStatus(ServiceControllerStatus.Stopped);
            }
            catch { }

            try
            {
                DWMService.Stop();
                DWMService.WaitForStatus(ServiceControllerStatus.Stopped);
            }
            catch { }

            btnStartBoth_Click(null, null);
        }

        private void tbDllName_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (TextBox)sender);
        }

        public void Edition(List<Control> list, Control control)
        {
            if (!list.Contains(control))
                list.Add(control);
        }

        private void tbDWMColorizationColor_TextChanged(object sender, EventArgs e)
        {
            try
            {
                panelDWMColorizationColor.BackColor = Color.FromArgb(int.Parse(tbDWMColorizationColor.Text, System.Globalization.NumberStyles.HexNumber));
            }
            catch
            {
                panelDWMColorizationColor.BackColor = Color.Red;
            }
            if (!IsUpdating)
                Edition(DWMEdited, (TextBox)sender);
        }

        private void panelDWMColorizationColor_Paint(object sender, PaintEventArgs e)
        {
        }

        public static void WriteResourceToFile(string resourceName, string fileName)
        {
            using (var resource = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
        }
        private void panelDWMColorizationColor_Click(object sender, EventArgs e)
        {

            ColorDialog c = new ColorDialog();
            c.Color = panelDWMColorizationColor.BackColor;
            c.FullOpen = true;
            if (c.ShowDialog() == DialogResult.OK)
            {
                tbDWMColorizationColor.Text = c.Color.ToArgb().ToString("X4");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog c = new ColorDialog();
            c.Color = panelColorBack.BackColor;
            c.FullOpen = true;
            if (c.ShowDialog() == DialogResult.OK)
            {
                panelColorBack.BackColor = c.Color;
            }
        }

        private void cbDWMCompos_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(DWMEdited, (CheckBox)sender);
        }

        private void cbLoadedBefore_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (CheckBox)sender);
        }

        private void cbThemeActive_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (CheckBox)sender);
        }

        private void tbCurrentTheme_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (TextBox)sender);
        }

        private void tbInstallVisualStyle_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (TextBox)sender);
        }

        private void cbDropShadow_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (CheckBox)sender);
        }

        private void cbFlatMenus_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(ThemesEdited, (CheckBox)sender);
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void cbHKCU_CheckedChanged(object sender, EventArgs e)
        {
            ThemesEdited.Clear();
            DWMEdited.Clear();
            CTEdited.Clear();
        }

        private void cbDefUser_CheckedChanged(object sender, EventArgs e)
        {
            ThemesEdited.Clear();
            DWMEdited.Clear();
            CTEdited.Clear();
        }

        private void cbHKLM_CheckedChanged(object sender, EventArgs e)
        {
            ThemesEdited.Clear();
            DWMEdited.Clear();
            CTEdited.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!File.Exists("desk2k.cpl"))
                WriteResourceToFile("ThemeTweaks.desk2k.cpl", "desk2k.cpl");
            if (!File.Exists("comctl2k.dll"))
                WriteResourceToFile("ThemeTweaks.comctl2k.dll", "comctl2k.dll");

            Process.Start("desk2k.cpl");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!File.Exists("desk2k.cpl"))
                WriteResourceToFile("ThemeTweaks.desk2k.cpl", "desk2k.cpl");
            if (!File.Exists("comctl2k.dll"))
                WriteResourceToFile("ThemeTweaks.comctl2k.dll", "comctl2k.dll");
            if (!File.Exists("themes2k.exe"))
                WriteResourceToFile("ThemeTweaks.themes2k.exe", "themes2k.exe");

            Process.Start("themes2k.exe");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPageCurrentTheme);
        }

        private void linkGoTabCurrentTheme_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tabControl1.SelectTab(tabPageCurrentTheme);
        }

        private void btnRefteshCTheme_Click(object sender, EventArgs e)
        {
            CTEdited.Clear();
        }

        private void btnEditCurrentTheme_Click(object sender, EventArgs e)
        {
            Process.Start("notepad", CurrentThemePath);
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {

        }

        private void btnAeroConvert_Click(object sender, EventArgs e)
        {
            tbPathCurrentTheme.Text = GetDefaultThemePath();
            tbColorStyleCurrentTheme.Text = "NormalColor";
        }

        private void btnClassicConvert_Click(object sender, EventArgs e)
        {
            btnThemingFalse_Click(null, null);
        }

        private void cbComposCurrentTheme_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(CTEdited, (CheckBox)sender);
        }

        private void cbTransparencyCurrentTheme_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(CTEdited, (CheckBox)sender);
        }

        private void tbPathCurrentTheme_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(CTEdited, (TextBox)sender);
        }

        private void tbColorStyleCurrentTheme_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(CTEdited, (TextBox)sender);
        }

        private void tbVisualStyleVersionCurrentTh_TextChanged(object sender, EventArgs e)
        {
            if (!IsUpdating)
                Edition(CTEdited, (TextBox)sender);
        }

        private void btnApplyCTheme_Click(object sender, EventArgs e)
        {
            try
            {
                string path = CurrentThemePath;// new FileInfo(Path.Combine(new FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, "new.theme")).FullName;

                if (rbConvertAero.Checked)
                {
                    cbThemeActive.Checked = true;
                    btnThemingTrue_Click(null, null);
                }
                else if (rbConvertClassic.Checked)
                {
                    btnThemingFalse_Click(null, null);
                    cbThemeActive.Checked = false;
                }

                rbNoConvert.Checked = true;

                if (cbRestartOnApplyCTheme.Checked)
                {
                    try
                    {
                        try
                        {
                            ThemesService.Stop();
                            ThemesService.WaitForStatus(ServiceControllerStatus.Stopped);
                        }
                        catch { }
                        cbLoadedBefore.Checked = false;

                        bool before = cbAutoRestart.Checked;
                        cbAutoRestart.Checked = false;

                        btnThemesApply_Click(null, null);

                        cbAutoRestart.Checked = before;
                    }
                    catch { }
                }


                string[] c = { };
                using (StreamReader reader = new StreamReader(CurrentThemePath))
                {
                    c = reader.ReadToEnd().Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                }

                for (int i = 0; i < c.Length; i++)
                {
                    if (c[i].StartsWith("[VisualStyles]"))
                        for (int x = i; x < c.Length; x++)
                        {
                            i = x;
                            if (c[i].Length <= 1)
                            {
                                break;
                            }
                            if (c[i].StartsWith("Path="))
                                c[i] = "Path=" + tbPathCurrentTheme.Text;

                            if (c[i].StartsWith("ColorStyle="))
                                c[i] = "ColorStyle=" + tbColorStyleCurrentTheme.Text;

                            if (c[i].StartsWith("VisualStyleVersion="))
                                c[i] = "VisualStyleVersion=" + tbVisualStyleVersionCurrentTh.Text;

                            if (c[i].StartsWith("Transparency="))
                            {
                                if (cbTransparencyCurrentTheme.Checked)
                                    c[i] = "Transparency=1";
                                else
                                    c[i] = "Transparency=0";
                            }

                            if (c[i].StartsWith("Composition="))
                            {
                                if (cbTransparencyCurrentTheme.Checked)
                                    c[i] = "Composition=1";
                                else
                                    c[i] = "Composition=0";
                            }
                        } 
                }

                using (StreamWriter writer = new StreamWriter(path))
                {
                    for (int i = 0; i < c.Length; i++)
                    {
                        writer.WriteLine(c[i]);
                    }
                }



                try
                {
                    CurrentThemePath = path;
                    

                    bool before = cbAutoRestart.Checked;
                    cbAutoRestart.Checked = false;

                    btnThemesApply_Click(null, null);

                    cbAutoRestart.Checked = before;
                }
                catch { }
                // ...
                CTEdited.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can't apply current theme settings");
            }

            if (cbRestartOnApplyCTheme.Checked)
                btnThemesRestart_Click(null, null);


        }

        private void cbRestartOnApplyCTheme_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnApplyVisualStyle_Click(object sender, EventArgs e)
        {
            int result = UxTheme.SetSystemVisualStyle(tbVisualStyleSet.Text, tbVisualColor.Text, tbVisualSize.Text, 0);
            lblVisualSetResult.Text = "Result: " + result;
        }

        private void btnThemingFalse_Click(object sender, EventArgs e)
        {
            int result = UxTheme.EnableTheming(false);
            lblVisualSetResult.Text = "Result: " + result;
        }

        private void btnThemingTrue_Click(object sender, EventArgs e)
        {
            int result = UxTheme.EnableTheming(true);
            lblVisualSetResult.Text = "Result: " + result;
        }

        private void rbNoConvert_CheckedChanged(object sender, EventArgs e)
        {
            if (CTEdited.Contains(tbPathCurrentTheme))
                CTEdited.Remove(tbPathCurrentTheme);

            if (CTEdited.Contains(tbColorStyleCurrentTheme))
                CTEdited.Remove(tbColorStyleCurrentTheme);
        }

        private void rbConvertAero_CheckedChanged(object sender, EventArgs e)
        {


            tbPathCurrentTheme.Text = GetDefaultThemePath();
            tbColorStyleCurrentTheme.Text = "NormalColor";
        }

        private void rbConvertClassic_CheckedChanged(object sender, EventArgs e)
        {


            tbPathCurrentTheme.Text = @"";
            tbColorStyleCurrentTheme.Text = "Normal";
        }
        public bool Kek = false;
        private void cbAeroWindowWPF_CheckedChanged(object sender, EventArgs e)
        {
            if (Kek)
            {
                Kek = false;
                return;
            }


            if (!File.Exists("AeroTest.exe"))
                WriteResourceToFile("ThemeTweaks.FTZ.exe", "AeroTest.exe");

            if (!File.Exists("FMUtils.KeyboardHook.dll"))
                WriteResourceToFile("ThemeTweaks.FMUtils.KeyboardHook.dll", "FMUtils.KeyboardHook.dll");

            Process.Start("AeroTest.exe");
            
            Kek = true;
            cbAeroWindowWPF.CheckState = CheckState.Indeterminate;
        }

        private void btnThemesDisableService_Click(object sender, EventArgs e)
        {
            if (ThemesService != null)
            {
                if (ThemesService.Status == ServiceControllerStatus.Running)
                    btnThemesStopStart_Click(null, null);

                ThemesService.WaitForStatus(ServiceControllerStatus.Stopped);

                ServiceHelper.ChangeStartMode(ThemesService, ServiceStartMode.Disabled);
                MessageBox.Show("Done.", "ThemeTweaks");
            }
            else
                MessageBox.Show("Themes Service not found");
        }

        private void btnThemesEnableService_Click(object sender, EventArgs e)
        {
            if (ThemesService != null)
            {
                ServiceHelper.ChangeStartMode(ThemesService, ServiceStartMode.Automatic);
                btnThemesStopStart_Click(null, null);
                MessageBox.Show("Done.", "ThemeTweaks");
            }
            else
                MessageBox.Show("Themes Service not found");
        }

        private void btnDWMServiceDisable_Click(object sender, EventArgs e)
        {
            if (DWMService != null)
            {
                if(DWMService.Status == ServiceControllerStatus.Running)
                    btnUxsmsStopStart_Click(null, null);

                DWMService.WaitForStatus(ServiceControllerStatus.Stopped);

                ServiceHelper.ChangeStartMode(DWMService, ServiceStartMode.Disabled);
                MessageBox.Show("Done.", "ThemeTweaks");
            }
            else
                MessageBox.Show("UXSMS Service not found");
        }

        private void btnDWMServiceEnable_Click(object sender, EventArgs e)
        {
            if (DWMService != null)
            {
                ServiceHelper.ChangeStartMode(DWMService, ServiceStartMode.Automatic);
                btnUxsmsStopStart_Click(null, null);
                MessageBox.Show("Done.", "ThemeTweaks");
            }
            else
                MessageBox.Show("UXSMS Service not found");
        }
    }
}