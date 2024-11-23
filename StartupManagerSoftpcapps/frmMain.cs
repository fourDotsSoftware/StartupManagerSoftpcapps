using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Xml;

namespace StartupManagerSoftpcapps
{
    public partial class frmMain : CustomForm
    {
        public static frmMain Instance = null;

        private DataGridView dgLastSelected = null;

        public frmMain()
        {
            InitializeComponent();

            Instance = this;

            //3Properties.Settings.Default.Password = "";

            dtLocalMachine.Columns.Add("icon", typeof(Bitmap));
            dtLocalMachine.Columns.Add("name");
            dtLocalMachine.Columns.Add("setting");
            dtLocalMachine.Columns.Add("comments");

            dtCurrentUser.Columns.Add("icon", typeof(Bitmap));
            dtCurrentUser.Columns.Add("name");
            dtCurrentUser.Columns.Add("setting");
            dtCurrentUser.Columns.Add("comments");

            dtRemovedLocalMachine.Columns.Add("icon", typeof(Bitmap));
            dtRemovedLocalMachine.Columns.Add("name");
            dtRemovedLocalMachine.Columns.Add("setting");
            dtRemovedLocalMachine.Columns.Add("comments");

            dtRemovedCurrentUser.Columns.Add("icon", typeof(Bitmap));
            dtRemovedCurrentUser.Columns.Add("name");
            dtRemovedCurrentUser.Columns.Add("setting");
            dtRemovedCurrentUser.Columns.Add("comments");

            dgLocalMachine.AutoGenerateColumns = false;
            dgCurrentUser.AutoGenerateColumns = false;

            dgLocalMachine.DataSource = dtLocalMachine;
            dgCurrentUser.DataSource = dtCurrentUser;

            dgRemovedLocalMachine.DataSource = dtRemovedLocalMachine;
            dgRemovedCurrentUser.DataSource = dtRemovedCurrentUser;

            dtLocalMachine.RowChanged += DtLocalMachine_RowChanged;
            dtCurrentUser.RowChanged += DtCurrentUser_RowChanged;
            dtLocalMachine.RowDeleted += DtLocalMachine_RowDeleted;
            dtCurrentUser.RowDeleted += DtCurrentUser_RowDeleted;
            dtLocalMachine.TableCleared += DtLocalMachine_TableCleared;
            dtCurrentUser.TableCleared += DtCurrentUser_TableCleared;
        }

        #region Total Settings

        private void DtCurrentUser_TableCleared(object sender, DataTableClearEventArgs e)
        {
            DtCurrentUser_RowChanged(null, null);
        }

        private void DtLocalMachine_TableCleared(object sender, DataTableClearEventArgs e)
        {
            DtLocalMachine_RowChanged(null, null);
        }

        private void DtCurrentUser_RowDeleted(object sender, DataRowChangeEventArgs e)
        {
            DtCurrentUser_RowChanged(null, null);
        }

        private void DtLocalMachine_RowDeleted(object sender, DataRowChangeEventArgs e)
        {
            DtLocalMachine_RowChanged(null, null);
        }

        private void DtCurrentUser_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            lblTotalCU.Text = TranslateHelper.Translate("Total") + " " + dtCurrentUser.Rows.Count.ToString() + " " + TranslateHelper.Translate("Settings");
        }

        private void DtLocalMachine_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            lblTotalLM.Text = TranslateHelper.Translate("Total") + " " + dtLocalMachine.Rows.Count.ToString() + " " + TranslateHelper.Translate("Settings");
        }

        #endregion

        public DataTable dtLocalMachine = new DataTable("table");
        public DataTable dtCurrentUser = new DataTable("table");
        public DataTable dtRemovedLocalMachine = new DataTable("table");
        public DataTable dtRemovedCurrentUser = new DataTable("table");


        private Bitmap EmptyBitmap = new Bitmap(16, 16);

        #region File Menu

        private void tsdbAddFile_ButtonClick(object sender, EventArgs e)
        {
            openFileDialog1.Filter = Module.OpenFilesFilter;
            openFileDialog1.Multiselect = true;

            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;

                    for (int k = 0; k < openFileDialog1.FileNames.Length; k++)
                    {
                        AddFile(openFileDialog1.FileNames[k]);
                        RecentFilesHelper.AddRecentFile(openFileDialog1.FileNames[k]);
                    }
                }
                finally
                {
                    this.Cursor = null;
                }
            }
        }

        private void tsdbAddFile_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                if (!System.IO.File.Exists(e.ClickedItem.Text))
                {
                    Module.ShowMessage("File does not exist !");
                    return;
                }

                this.Cursor = Cursors.WaitCursor;

                AddFile(e.ClickedItem.Text);
                RecentFilesHelper.AddRecentFile(e.ClickedItem.Text);

            }
            finally
            {
                this.Cursor = null;
            }
        }

        private void tsbRemove_Click(object sender, EventArgs e)
        {
            DialogResult dres = Module.ShowQuestionDialog(
                TranslateHelper.Translate("Are you sure that you want to Remove Setting ? You will be able to include it again from the 'Removed' tab !"),
                TranslateHelper.Translate("Remove Setting ?")
                );

            if (dres != DialogResult.Yes)
            {
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                List<DataGridViewRow> rows = new List<DataGridViewRow>();

                if (tabControl1.SelectedIndex == 0)
                {
                    for (int k = 0; k < dgLocalMachine.SelectedRows.Count; k++)
                    {
                        rows.Add(dgLocalMachine.SelectedRows[k]);
                    }
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    for (int k = 0; k < dgCurrentUser.SelectedRows.Count; k++)
                    {
                        rows.Add(dgCurrentUser.SelectedRows[k]);
                    }
                }
                else
                {
                    for (int k = 0; k < dgLastSelected.SelectedRows.Count; k++)
                    {
                        rows.Add(dgLastSelected.SelectedRows[k]);
                    }
                }

                for (int k = 0; k < rows.Count; k++)
                {
                    DataRowView drv = (DataRowView)rows[k].DataBoundItem;

                    if (tabControl1.SelectedIndex == 0)
                    {
                        DataRow dr = dtRemovedLocalMachine.NewRow();
                        dr["icon"] = drv.Row["icon"];
                        dr["name"] = drv.Row["name"];
                        dr["setting"] = drv.Row["setting"];
                        dr["comments"] = drv.Row["comments"];
                        dtRemovedLocalMachine.Rows.Add(dr);
                    }
                    if (tabControl1.SelectedIndex == 1)
                    {
                        DataRow dr = dtRemovedCurrentUser.NewRow();
                        dr["icon"] = drv.Row["icon"];
                        dr["name"] = drv.Row["name"];
                        dr["setting"] = drv.Row["setting"];
                        dr["comments"] = drv.Row["comments"];
                        dtRemovedCurrentUser.Rows.Add(dr);
                    }
                }

                for (int k = 0; k < rows.Count; k++)
                {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        dgLocalMachine.Rows.Remove(rows[k]);
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        dgCurrentUser.Rows.Remove(rows[k]);
                    }
                    else
                    {
                        dgLastSelected.Rows.Remove(rows[k]);
                    }
                }
            }
            finally
            {
                this.Cursor = null;
            }
        }
        public void ImportList(string listfilepath)
        {
            try
            {

                using (StreamReader sr = new StreamReader(listfilepath, Encoding.Default, true))
                {
                    string line = null;

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.StartsWith("#"))
                        {
                            continue;
                        }

                        string name = "";
                        string command = "";
                        string comments = "";

                        int spos = 0;
                        int epos = 0;

                        for (int k = 0; k < line.Length; k++)
                        {
                            bool indquotes = false;
                            bool insquotes = false;
                            bool justclosedquotes = false;

                            int sdq = -1;
                            int ssq = -1;

                            if (!insquotes && !indquotes)
                            {
                                if (line[k] == '\"')
                                {
                                    indquotes = true;
                                    sdq = k;
                                }
                            }
                            if (!insquotes && indquotes)
                            {
                                if (line[k] == '\"')
                                {
                                    indquotes = false;
                                    justclosedquotes = true;
                                }
                            }

                            if (!insquotes && !indquotes)
                            {
                                if (line[k] == '\'')
                                {
                                    insquotes = true;
                                    ssq = k;
                                }
                            }
                            if (insquotes && !indquotes)
                            {
                                if (line[k] == '\'')
                                {
                                    insquotes = false;
                                    justclosedquotes = true;
                                }
                            }

                            if (line[k] == ',' && !indquotes && !insquotes)
                            {
                                epos = k - 1;

                                //01234
                                if (command == string.Empty)
                                {
                                    command = line.Substring(spos, epos - spos + 1);
                                }
                                else if (command != string.Empty)
                                {
                                    name = command;
                                    command = line.Substring(spos, epos - spos + 1);
                                }

                                spos = k + 1;
                            }

                            if (k == line.Length - 1)
                            {
                                epos = line.Length - 1;

                                if (command == string.Empty)
                                {
                                    command = line.Substring(spos, epos - spos + 1);
                                }
                                else if (command != string.Empty && name == string.Empty)
                                {
                                    name = command;
                                    command = line.Substring(spos, epos - spos + 1);
                                }
                                else if (command != string.Empty && name != string.Empty)
                                {
                                    comments = line.Substring(spos, epos - spos + 1);
                                }
                            }

                        }

                        if (command.Length > 2 && (command.StartsWith("\"") && command.EndsWith("\"")))
                        {
                            command = command.Substring(1, command.Length - 2);
                        }

                        if (command.Length > 2 && (command.StartsWith("'") && command.EndsWith("'")))
                        {
                            command = command.Substring(1, command.Length - 2);
                        }

                        if (name.Length > 2 && (name.StartsWith("\"") && name.EndsWith("\"")))
                        {
                            name = name.Substring(1, name.Length - 2);
                        }

                        if (name.Length > 2 && (name.StartsWith("'") && name.EndsWith("'")))
                        {
                            name = name.Substring(1, name.Length - 2);
                        }

                        if (comments.Length > 2 && (comments.StartsWith("\"") && comments.EndsWith("\"")))
                        {
                            comments = comments.Substring(1, comments.Length - 2);
                        }

                        if (comments.Length > 2 && (comments.StartsWith("'") && comments.EndsWith("'")))
                        {
                            comments = comments.Substring(1, comments.Length - 2);
                        }

                        AddFile(command, name, comments);
                    }
                }
            }
            catch (Exception ex)
            {
                throw (ex);
                //throw new Exception("Error could not parse Import File");
            }
            finally
            {

            }
        }

        private void tsdbImportList_ButtonClick(object sender, EventArgs e)
        {
            //SilentAddErr = "";

            openFileDialog1.Filter = "Text Files (*.txt),CSV Files (*.csv)|*.txt;*.csv|Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ImportList(openFileDialog1.FileName);
                RecentFilesHelper.ImportListRecent(openFileDialog1.FileName);
                /*
                if (SilentAddErr != string.Empty)
                {
                    frmMessage f = new frmMessage();
                    f.txtMsg.Text = SilentAddErr;
                    f.ShowDialog();

                }*/
            }
        }

        private void tsdbImportList_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (!System.IO.File.Exists(e.ClickedItem.Text))
            {
                Module.ShowMessage("File does not exist !");
                return;
            }

            ImportList(e.ClickedItem.Text);
            RecentFilesHelper.ImportListRecent(e.ClickedItem.Text);

        }

        #endregion

        #region Share

        private void tsiFacebook_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareFacebook();
        }

        private void tsiTwitter_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareTwitter();
        }

        private void tsiGooglePlus_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareGooglePlus();
        }

        private void tsiLinkedIn_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareLinkedIn();
        }

        private void tsiEmail_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareEmail();
        }

        #endregion

        public bool AddFile(string filepath)
        {
            return AddFile(filepath, "", "");
        }


        public bool AddFile(string command, string name, string comments)
        {
            try
            {
                DataRow dr = null;

                if (tabControl1.SelectedIndex == 0)
                {
                    dr = dtLocalMachine.NewRow();
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    dr = dtCurrentUser.NewRow();
                }
                else
                {
                    return true;
                }

                Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(command);

                if (bmp == null)
                {
                    bmp = EmptyBitmap;
                }

                string n1 = "";

                try
                {
                    n1 = System.IO.Path.GetFileNameWithoutExtension(command);
                }
                catch { }

                dr["icon"] = bmp;
                dr["name"] = name == string.Empty ? n1 : name;
                dr["setting"] = command;
                dr["comments"] = comments;

                if (tabControl1.SelectedIndex == 0)
                {
                    dtLocalMachine.Rows.Add(dr);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    dtCurrentUser.Rows.Add(dr);
                }

                return true;
            }
            finally
            {

            }
        }

        bool FreeForPersonalUse = false;
        bool FreeForPersonalAndCommercialUse = true;

        private void SetTitle()
        {
            string str = "";
                        
            if (FreeForPersonalUse)
            {
                str += " - " + TranslateHelper.Translate("Free for Personal Use Only - Please Donate !");
            }
            else if (FreeForPersonalAndCommercialUse)
            {
                str += " - " + TranslateHelper.Translate("Free for Personal and Commercial Use - Please Donate !");
            }

            this.Text = Module.ApplicationTitle + str.ToUpper();
        }
        public void SetupOnLoad()
        {
            dtLocalMachine.CaseSensitive = false;

            //3this.Icon = Properties.Resources.pdf_compress_48;

            this.Text = Module.ApplicationTitle;

            SetTitle();

            //this.Width = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            //this.Left = 0;
            //12AddLanguageMenuItems();

            AdjustSizeLocation();

            RecentFilesHelper.FillMenuRecentFile();
            RecentFilesHelper.FillMenuRecentImportList();

            ResizeControls();

            doNotImportDuplicatesToolStripMenuItem.Checked = Properties.Settings.Default.DoNotImportDuplicates;

            //12checkForNewVersionEachWeekToolStripMenuItem.Checked = Properties.Settings.Default.CheckWeek;            

            FillSettings();
        }

        private void FillSettings()
        {
            StartupSettingsManager s1 = new StartupSettingsManager(StartupSettingsManager.StartupSettingsType.HKLM);

            string[] s1settings = s1.GetSettingNames();

            string[] s1values = s1.GetSettingValues();

            StartupSettingsManager s2 = new StartupSettingsManager(StartupSettingsManager.StartupSettingsType.HKCU);

            string[] s2settings = s2.GetSettingNames();

            string[] s2values = s2.GetSettingValues();

            dtLocalMachine.Rows.Clear();

            for (int k = 0; k < s1settings.Length; k++)
            {
                DataRow dr = dtLocalMachine.NewRow();

                Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(s1values[k]);

                if (bmp == null)
                {
                    bmp = EmptyBitmap;
                }

                dr["icon"] = bmp;
                dr["name"] = s1settings[k];
                dr["setting"] = s1values[k];
                dr["comments"] = "";

                dtLocalMachine.Rows.Add(dr);

            }

            for (int k = 0; k < s2settings.Length; k++)
            {
                DataRow dr = dtCurrentUser.NewRow();

                Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(s2values[k]);

                if (bmp == null)
                {
                    bmp = EmptyBitmap;
                }

                dr["icon"] = bmp;
                dr["name"] = s2settings[k];
                dr["setting"] = s2values[k];
                dr["comments"] = "";

                dtCurrentUser.Rows.Add(dr);

            }

            SettingsFileManager sfm = new SettingsFileManager();

            sfm.LoadSettings();

        }
        private void AdjustSizeLocation()
        {
            if (Properties.Settings.Default.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {

                if (Properties.Settings.Default.Width == -1)
                {
                    this.CenterToScreen();
                    return;
                }
                else
                {
                    this.Width = Properties.Settings.Default.Width;
                }
                if (Properties.Settings.Default.Height != 0)
                {
                    this.Height = Properties.Settings.Default.Height;
                }

                if (Properties.Settings.Default.Left != -1)
                {
                    this.Left = Properties.Settings.Default.Left;
                }

                if (Properties.Settings.Default.Top != -1)
                {
                    this.Top = Properties.Settings.Default.Top;
                }

                if (this.Width < 300)
                {
                    this.Width = 300;
                }

                if (this.Height < 300)
                {
                    this.Height = 300;
                }

                if (this.Left < 0)
                {
                    this.Left = 0;
                }

                if (this.Top < 0)
                {
                    this.Top = 0;
                }
            }

        }

        private void SaveSizeLocation()
        {
            Properties.Settings.Default.Maximized = (this.WindowState == FormWindowState.Maximized);

            if (this.WindowState == System.Windows.Forms.FormWindowState.Minimized) return;

            if (this.WindowState == System.Windows.Forms.FormWindowState.Maximized)
            {
                Properties.Settings.Default.Save();
                return;
            }

            Properties.Settings.Default.Left = this.Left;
            Properties.Settings.Default.Top = this.Top;
            Properties.Settings.Default.Width = this.Width;
            Properties.Settings.Default.Height = this.Height;
            Properties.Settings.Default.Save();

        }
        /*12
        #region Localization

        private void AddLanguageMenuItems()
        {
            for (int k = 0; k < frmLanguage.LangCodes.Count; k++)
            {
                ToolStripMenuItem ti = new ToolStripMenuItem();
                ti.Text = frmLanguage.LangDesc[k];
                ti.Tag = frmLanguage.LangCodes[k];
                ti.Image = frmLanguage.LangImg[k];

                if (Properties.Settings.Default.Language == frmLanguage.LangCodes[k])
                {
                    ti.Checked = true;
                }

                ti.Click += new EventHandler(tiLang_Click);

                if (k < 25)
                {
                    languages1ToolStripMenuItem.DropDownItems.Add(ti);
                }
                else
                {
                    languages2ToolStripMenuItem.DropDownItems.Add(ti);
                }

                //languageToolStripMenuItem.DropDownItems.Add(ti);
            }
        }

        void tiLang_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ti = (ToolStripMenuItem)sender;
            string langcode = ti.Tag.ToString();
            ChangeLanguage(langcode);

            //for (int k = 0; k < languageToolStripMenuItem.DropDownItems.Count; k++)
            for (int k = 0; k < languages1ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages1ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }

            for (int k = 0; k < languages2ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages2ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }
        }

        private bool InChangeLanguage = false;

        private void ChangeLanguage(string language_code)
        {
            try
            {
                InChangeLanguage = true;

                Properties.Settings.Default.Language = language_code;
                frmLanguage.SetLanguage();

                Module.ShowMessage("Please restart the application !");
                Application.Exit();

                return;

                bool maximized = (this.WindowState == FormWindowState.Maximized);
                this.WindowState = FormWindowState.Normal;

                /*
                RegistryKey key = Registry.CurrentUser;
                RegistryKey key2 = Registry.CurrentUser;

                try
                {
                    key = key.OpenSubKey("Software\\softpcapps Software", true);

                    if (key == null)
                    {
                        key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\softpcapps Software");
                    }

                    key2 = key.OpenSubKey(frmLanguage.RegKeyName, true);

                    if (key2 == null)
                    {
                        key2 = key.CreateSubKey(frmLanguage.RegKeyName);
                    }

                    key = key2;

                    //key.SetValue("Language", language_code);
                    key.SetValue("Menu Item Caption", TranslateHelper.Translate("Change PDF Properties"));
                }
                catch (Exception ex)
                {
                    Module.ShowError(ex);
                    return;
                }
                finally
                {
                    key.Close();
                    key2.Close();
                }
                */
        //1SaveSizeLocation();

        //3SavePositionSize();\
        /*12

        this.Controls.Clear();

        InitializeComponent();

        SetupOnLoad();

        if (maximized)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        this.ResumeLayout(true);
    }
    finally
    {
        InChangeLanguage = false;
    }
}

#endregion
*/

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {


                SetupOnLoad();

                /*12
                if (Properties.Settings.Default.CheckWeek)
                {
                    UpdateHelper.InitializeCheckVersionWeek();
                }
                */
            }
            finally
            {
                this.ResumeLayout();
            }
        }

        #region Help

        private void helpGuideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(Application.StartupPath + "\\Video Cutter Joiner Expert - User's Manual.chm");
            System.Diagnostics.Process.Start(Module.HelpURL);
        }

        private void pleaseDonateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://softpcapps.com/donate.php");
        }

        private void dotsSoftwarePRODUCTCATALOGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://softpcapps.com/downloads/4dots-Software-PRODUCT-CATALOG.pdf");
        }        
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            frmAbout f = new frmAbout();
            f.ShowDialog();
        }

        private void tiHelpFeedback_Click(object sender, EventArgs e)
        {
            /*
            frmUninstallQuestionnaire f = new frmUninstallQuestionnaire(false);
            f.ShowDialog();
            */

            System.Diagnostics.Process.Start("https://softpcapps.com/support/bugfeature.php?app=" + System.Web.HttpUtility.UrlEncode(Module.ShortApplicationTitle));
        }

        private void followUsOnTwitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.twitter.com/4dotsSoftware");
        }

        private void visit4dotsSoftwareWebsiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://softpcapps.com");
        }

        private void checkForNewVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateHelper.CheckVersion(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch { }
        }

        #endregion

        public void SaveOptions()
        {
            Properties.Settings.Default.CheckWeek = checkForNewVersionEachWeekToolStripMenuItem.Checked;

            Properties.Settings.Default.Save();
        }
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason != CloseReason.WindowsShutDown)
            {
                DialogResult dres = Module.ShowQuestionDialog(TranslateHelper.Translate("Are you sure that you want to exit") + " " + Module.ApplicationName + " ?"
                    , TranslateHelper.Translate("Exit") + " " + Module.ApplicationName + " ?");

                if (dres != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
            }

            SettingsFileManager sfm = new SettingsFileManager();
            sfm.SaveSettings();

            SaveOptions();
            SaveSizeLocation();

            Properties.Settings.Default.Save();
        }

        #region Edit Menu

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                for (int k = 0; k < dgLocalMachine.Rows.Count; k++)
                {
                    dgLocalMachine.Rows[k].Selected = true;
                }
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                for (int k = 0; k < dgCurrentUser.Rows.Count; k++)
                {
                    dgCurrentUser.Rows[k].Selected = true;
                }
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                for (int k = 0; k < dgLastSelected.Rows.Count; k++)
                {
                    dgLastSelected.Rows[k].Selected = true;
                }
            }
        }

        private void seelctNoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                for (int k = 0; k < dgLocalMachine.Rows.Count; k++)
                {
                    dgLocalMachine.Rows[k].Selected = false;
                }
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                for (int k = 0; k < dgCurrentUser.Rows.Count; k++)
                {
                    dgCurrentUser.Rows[k].Selected = false;
                }
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                for (int k = 0; k < dgLastSelected.Rows.Count; k++)
                {
                    dgLastSelected.Rows[k].Selected = false;
                }
            }
        }

        private void invertSelectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                for (int k = 0; k < dgLocalMachine.Rows.Count; k++)
                {
                    dgLocalMachine.Rows[k].Selected = !dgLocalMachine.Rows[k].Selected;
                }
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                for (int k = 0; k < dgCurrentUser.Rows.Count; k++)
                {
                    dgCurrentUser.Rows[k].Selected = !dgCurrentUser.Rows[k].Selected;
                }
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                for (int k = 0; k < dgLastSelected.Rows.Count; k++)
                {
                    dgLastSelected.Rows[k].Selected = !dgLastSelected.Rows[k].Selected;
                }
            }
        }

        #endregion

        #region Grid Context menu

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgLocalMachine.CurrentRow == null) return;

            DataRowView drv = (DataRowView)dgLocalMachine.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            System.Diagnostics.Process.Start(filepath);
        }

        private void exploreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgLastSelected.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = SettingsParser.GetExeFilepath(dr["setting"].ToString());

            if (filepath == string.Empty) return;

            string args = string.Format("/e, /select, \"{0}\"", filepath);

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "explorer";
            info.UseShellExecute = true;
            info.Arguments = args;
            Process.Start(info);
        }

        private void copyFullFilePathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgLocalMachine.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["setting"].ToString();

            if (filepath == string.Empty) return;

            Clipboard.Clear();

            Clipboard.SetText(filepath);
        }

        private void cmsFiles_Opening(object sender, CancelEventArgs e)
        {            
            Point p = dgLastSelected.PointToClient(new Point(Control.MousePosition.X, Control.MousePosition.Y));

            DataGridView.HitTestInfo hit = dgLastSelected.HitTest(p.X, p.Y);

            if (hit.Type == DataGridViewHitTestType.Cell)
            {
                dgLastSelected.CurrentCell = dgLastSelected.Rows[hit.RowIndex].Cells[hit.ColumnIndex];
            }

            if (dgLastSelected.CurrentRow == null)
            {
                e.Cancel = true;

                return;
            }

            DataRowView drv = (DataRowView)dgLastSelected.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            includeToolStripMenuItem.Visible = tabControl1.SelectedIndex == 2;            
        }
        #endregion                

        #region Drag and Drop

        private void dgFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dgFiles_DragOver(object sender, DragEventArgs e)
        {
            if ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy)
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dgFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                string[] filez = (string[])e.Data.GetData(DataFormats.FileDrop);

                for (int k = 0; k < filez.Length; k++)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;

                        if (System.IO.File.Exists(filez[k]))
                        {
                            AddFile(filez[k]);
                        }
                        else if (System.IO.Directory.Exists(filez[k]))
                        {

                        }
                    }
                    finally
                    {
                        this.Cursor = null;
                    }
                }
            }
        }

        #endregion                                                             
        private void youtubeChannelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*
            FileShower fs = new FileShower(@"c:\1\03\con.a.jpg");
            fs.Show();

            Module.ShowError(fs.Err);
            */

            System.Diagnostics.Process.Start("https://www.youtube.com/channel/UCovA-lld9Q79l08K-V1QEng");
        }
        private void importListFromExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel Files (*.xls;*.xlsx;*.xlt)|*.xls;*.xlsx;*.xlt";
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelImporter xl = new ExcelImporter();
                xl.ImportListExcel(openFileDialog1.FileName);
            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int h = tabPage3.Height;

            if (tabControl1.SelectedIndex == 2)
            {
                niceLine1.Top = 0;
                dgRemovedLocalMachine.Top = niceLine1.Bottom;
                dgRemovedLocalMachine.Height = h / 2 - 2 * niceLine1.Height;

                niceLine2.Top = dgRemovedLocalMachine.Bottom;
                dgRemovedCurrentUser.Top = niceLine2.Bottom;
                dgRemovedCurrentUser.Height = h / 2 - 2 * niceLine1.Height;

            }
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            tabControl1_SelectedIndexChanged(null, null);
        }

        private void dgLocalMachine_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = dgLocalMachine;
        }

        private void dgCurrentUser_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = dgCurrentUser;
        }

        private void dgRemovedLocalMachine_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = dgRemovedLocalMachine;
        }

        private void dgRemovedCurrentUser_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = dgRemovedCurrentUser;
        }

        private void dgRemovedCurrentUser_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void dgRemovedLocalMachine_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void tsbApply2_Click(object sender, EventArgs e)
        {
            dgCurrentUser.EndEdit();
            dgLocalMachine.EndEdit();

            StartupSettingsManager sm = null;

            if (tabControl1.SelectedIndex == 0)
            {
                DialogResult dres = Module.ShowQuestionDialog(
               TranslateHelper.Translate("Are you sure that you want to Apply Local Machine Startup Settings ? (not Current User ones)")
               , TranslateHelper.Translate("Apply Local Machine Startup Settings ?"));

                if (dres != DialogResult.Yes) return;

                sm = new StartupSettingsManager(StartupSettingsManager.StartupSettingsType.HKLM, dtLocalMachine);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                DialogResult dres = Module.ShowQuestionDialog(
               TranslateHelper.Translate("Are you sure that you want to Apply Current User Startup Settings ? (not Local Machine ones)")
               , TranslateHelper.Translate("Apply Current User Startup Settings ?"));

                if (dres != DialogResult.Yes) return;

                sm = new StartupSettingsManager(StartupSettingsManager.StartupSettingsType.HKCU, dtCurrentUser);
            }

            string val = sm.ValidateSettings();

            if (val == string.Empty)
            {
                sm.ApplySettings();

                Module.ShowMessage(TranslateHelper.Translate("Operation completed successfully !"));
            }
            else
            {
                Module.ShowMessage(val);
            }
        }

        private void dgRemovedLocalMachine_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = (DataGridView)sender;

            if (e.ColumnIndex==colInclude.DisplayIndex)
            {
                List<DataGridViewRow> drs = new List<DataGridViewRow>();

                for (int k=0;k<dgRemovedLocalMachine.SelectedRows.Count;k++)
                {
                    drs.Add(dgRemovedLocalMachine.SelectedRows[k]);
                
                    DataRowView drv = (DataRowView)dgRemovedLocalMachine.SelectedRows[k].DataBoundItem;
                    DataRow dr = drv.Row;

                    DataRow dr1 = dtLocalMachine.NewRow();

                    for (int m=0;m<dr.Table.Columns.Count;m++)
                    {
                        dr1[m] = dr[m];
                    }

                    dtLocalMachine.Rows.Add(dr1);
                }

                for (int k=0;k<drs.Count;k++)
                {
                    dgRemovedLocalMachine.Rows.Remove(drs[k]);
                }
            }
        }

        private void dgRemovedCurrentUser_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = (DataGridView)sender;

            if (e.ColumnIndex == colInclude.DisplayIndex)
            {
                List<DataGridViewRow> drs = new List<DataGridViewRow>();

                for (int k = 0; k < dgRemovedCurrentUser.SelectedRows.Count; k++)
                {
                    drs.Add(dgRemovedCurrentUser.SelectedRows[k]);

                    DataRowView drv = (DataRowView)dgRemovedCurrentUser.SelectedRows[k].DataBoundItem;
                    DataRow dr = drv.Row;

                    DataRow dr1 = dtCurrentUser.NewRow();

                    for (int m = 0; m < dr.Table.Columns.Count; m++)
                    {
                        dr1[m] = dr[m];
                    }

                    dtCurrentUser.Rows.Add(dr1);
                }

                for (int k = 0; k < drs.Count; k++)
                {
                    dgRemovedCurrentUser.Rows.Remove(drs[k]);
                }
            }
        }

        private void loadProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            
            ofd.Filter = "Startup Manager Softpcapps Project (*.smp)|*.smp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ProjectManager pr = new ProjectManager(ProjectManager.ProjectType.HKLM,ofd.FileName);
                pr.Load();
            }
        }

        private void DatagridsEndEdit()
        {
            dgLocalMachine.EndEdit();
            dgRemovedLocalMachine.EndEdit();
            dgCurrentUser.EndEdit();
            dgRemovedCurrentUser.EndEdit();
        }
        private void saveProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DatagridsEndEdit();

            SaveFileDialog ofd = new SaveFileDialog();

            ofd.Filter = "Startup Manager Softpcapps Project (*.smp)|*.smp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ProjectManager pr = new ProjectManager(ProjectManager.ProjectType.HKLM, ofd.FileName);
                pr.Save(dtLocalMachine, dtRemovedLocalMachine);
            }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Startup Manager Softpcapps Project (*.smp)|*.smp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ProjectManager pr = new ProjectManager(ProjectManager.ProjectType.HKCU, ofd.FileName);
                pr.Load();
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            DatagridsEndEdit();

            SaveFileDialog ofd = new SaveFileDialog();

            ofd.Filter = "Startup Manager Softpcapps Project (*.smp)|*.smp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ProjectManager pr = new ProjectManager(ProjectManager.ProjectType.HKCU, ofd.FileName);
                pr.Save(dtCurrentUser, dtRemovedCurrentUser);
            }
        }

        private void doNotImportDuplicatesToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DoNotImportDuplicates = doNotImportDuplicatesToolStripMenuItem.Checked;
        }

        private void tsbClear_Click_1(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0 || tabControl1.SelectedIndex == 1)
            {
                DialogResult dres = Module.ShowQuestionDialog(TranslateHelper.Translate("Are you sure that you want to Clear current Settings ? You will be able to include them again from the 'Removed' Tab Page"),
                    TranslateHelper.Translate("Clear Settings ?"));

                if (dres == DialogResult.Yes)
                {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        for (int k = 0; k < dtLocalMachine.Rows.Count; k++)
                        {
                            DataRow dr = dtRemovedLocalMachine.NewRow();

                            for (int m = 0; m < dtLocalMachine.Columns.Count; m++)
                            {
                                dr[m] = dtLocalMachine.Rows[k][m];
                            }

                            dtRemovedLocalMachine.Rows.Add(dr);
                        }

                        dtLocalMachine.Rows.Clear();
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        for (int k = 0; k < dtCurrentUser.Rows.Count; k++)
                        {
                            DataRow dr = dtRemovedCurrentUser.NewRow();

                            for (int m = 0; m < dtCurrentUser.Columns.Count; m++)
                            {
                                dr[m] = dtCurrentUser.Rows[k][m];
                            }

                            dtRemovedCurrentUser.Rows.Add(dr);
                        }

                        dtCurrentUser.Rows.Clear();
                    }
                }
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                DialogResult dres = Module.ShowQuestionDialog(TranslateHelper.Translate("Are you sure that you want to Clear current Settings ? You will NOT be able to include them again from the 'Removed' Tab Page"),
                    TranslateHelper.Translate("Clear Settings ?"));

                if (dres==DialogResult.Yes)
                {
                    dtRemovedCurrentUser.Rows.Clear();
                    dtRemovedLocalMachine.Rows.Clear();
                }
            }

        }

        private void enterFileListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string txt = "";            

            frmMultipleFiles f = new frmMultipleFiles(false, txt);

            if (f.ShowDialog() == DialogResult.OK)
            {
                for (int k = 0; k < f.txtFiles.Lines.Length; k++)
                {
                    AddFile(f.txtFiles.Lines[k].Trim());
                }
            }

        }

        private void saveCurrentFileListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();                        

            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Text Files (*.txt)|*.txt";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName, false, Encoding.Default))
                {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        for (int k = 0; k < dtLocalMachine.Rows.Count; k++)
                        {
                            sw.WriteLine(dtLocalMachine.Rows[k]["setting"].ToString());
                        }
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        for (int k = 0; k < dtCurrentUser.Rows.Count; k++)
                        {
                            sw.WriteLine(dtCurrentUser.Rows[k]["setting"].ToString());
                        }
                    }
                }
            }

        }

        private void tsbAddRow_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex==0)
            {
                DataRow dr = dtLocalMachine.NewRow();
                dr["icon"] = EmptyBitmap;
                dtLocalMachine.Rows.Add(dr);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                DataRow dr = dtCurrentUser.NewRow();
                dr["icon"] = EmptyBitmap;
                dtCurrentUser.Rows.Add(dr);
            }
        }

        private void removeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dgLastSelected!=null && dgLastSelected.SelectedRows.Count>0)
            {
                tsbRemove_Click(null, null);
            }            
        }

        private void includeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 2)
            {
                if (dgLastSelected == dgRemovedLocalMachine)
                {
                    List<DataGridViewRow> drs = new List<DataGridViewRow>();

                    for (int k = 0; k < dgRemovedLocalMachine.SelectedRows.Count; k++)
                    {
                        drs.Add(dgRemovedLocalMachine.SelectedRows[k]);

                        DataRowView drv = (DataRowView)dgRemovedLocalMachine.SelectedRows[k].DataBoundItem;
                        DataRow dr = drv.Row;

                        DataRow dr1 = dtLocalMachine.NewRow();

                        for (int m = 0; m < dr.Table.Columns.Count; m++)
                        {
                            dr1[m] = dr[m];
                        }

                        dtLocalMachine.Rows.Add(dr1);
                    }

                    for (int k = 0; k < drs.Count; k++)
                    {
                        dgRemovedLocalMachine.Rows.Remove(drs[k]);
                    }
                }
                else if (dgLastSelected == dgRemovedCurrentUser)
                {
                    List<DataGridViewRow> drs = new List<DataGridViewRow>();

                    for (int k = 0; k < dgRemovedCurrentUser.SelectedRows.Count; k++)
                    {
                        drs.Add(dgRemovedCurrentUser.SelectedRows[k]);

                        DataRowView drv = (DataRowView)dgRemovedCurrentUser.SelectedRows[k].DataBoundItem;
                        DataRow dr = drv.Row;

                        DataRow dr1 = dtCurrentUser.NewRow();

                        for (int m = 0; m < dr.Table.Columns.Count; m++)
                        {
                            dr1[m] = dr[m];
                        }

                        dtCurrentUser.Rows.Add(dr1);
                    }

                    for (int k = 0; k < drs.Count; k++)
                    {
                        dgRemovedCurrentUser.Rows.Remove(drs[k]);
                    }
                }
            }
        }

        private void dgLocalMachine_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = (DataGridView)sender;
        }

        private void dgCurrentUser_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgLastSelected = (DataGridView)sender;
        }

        private void dgLocalMachine_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void dgLocalMachine_MouseDown(object sender, MouseEventArgs e)
        {
            dgLastSelected = (DataGridView)sender;
        }
    }
}
