using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.Windows.Forms;
using System.Drawing;
using System.Security;

namespace StartupManagerSoftpcapps
{
    public class ProjectManager
    {
        public enum ProjectType
        {
            HKLM,HKCU
        }

        private ProjectType _ProjectType;
        private string _Filepath;
        private Bitmap EmptyBitmap = new Bitmap(16, 16);
        public ProjectManager(ProjectType projectType,string filepath)
        {
            _ProjectType = projectType;
            _Filepath = filepath;
        }

        public void Load()
        {
            if (!System.IO.File.Exists(_Filepath))
            {
                throw new Exception("File does not exist !");                
            }

            string xml = System.IO.File.ReadAllText(_Filepath);

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(xml);
            }
            catch
            {
                throw new Exception("Invalid Project !");
            }

            DialogResult dres = Module.ShowQuestionDialog(
                TranslateHelper.Translate("Would you like to Import the Project or load an entirely new Project and clear the existing one ?"),
                TranslateHelper.Translate("Import Project ?"));

            if (dres == DialogResult.Cancel) return;


            if (_ProjectType == ProjectType.HKLM)
            {
                if (dres == DialogResult.No)
                {
                    frmMain.Instance.dtLocalMachine.Clear();
                    frmMain.Instance.dtRemovedLocalMachine.Clear();
                }

                XmlNode no = doc.SelectSingleNode("//LocalMachine");
                XmlNodeList nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string name = nolos[m].Attributes.GetNamedItem("Name").Value;
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    bool found = false;

                    if (Properties.Settings.Default.DoNotImportDuplicates)
                    {
                        for (int k = 0; k < frmMain.Instance.dtLocalMachine.Rows.Count; k++)
                        {
                            if (frmMain.Instance.dtLocalMachine.Rows[k]["name"].ToString().ToLower() == name.ToLower())
                            {
                                found = true;
                                if (frmMain.Instance.dtLocalMachine.Rows[k]["comments"].ToString() == string.Empty)
                                {
                                    frmMain.Instance.dtLocalMachine.Rows[k]["comments"] = comments;
                                }
                            }
                        }
                    }

                    if (!found || !Properties.Settings.Default.DoNotImportDuplicates)
                    {
                        DataRow dr = frmMain.Instance.dtLocalMachine.NewRow();

                        Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(command);

                        if (bmp == null)
                        {
                            bmp = EmptyBitmap;
                        }

                        dr["icon"] = bmp;
                        dr["name"] = name;
                        dr["setting"] = command;
                        dr["comments"] = comments;

                        frmMain.Instance.dtLocalMachine.Rows.Add(dr);
                    }
                }
                
                no = doc.SelectSingleNode("//RemovedLocalMachine");
                nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string name = nolos[m].Attributes.GetNamedItem("Name").Value;
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(command);

                    if (bmp == null)
                    {
                        bmp = EmptyBitmap;
                    }

                    bool found = false;

                    if (Properties.Settings.Default.DoNotImportDuplicates)
                    {
                        for (int k = 0; k < frmMain.Instance.dtRemovedLocalMachine.Rows.Count; k++)
                        {
                            if (frmMain.Instance.dtRemovedLocalMachine.Rows[k]["name"].ToString().ToLower() == name.ToLower())
                            {
                                found = true;

                                if (frmMain.Instance.dtRemovedLocalMachine.Rows[k]["comments"].ToString() == string.Empty)
                                {
                                    frmMain.Instance.dtRemovedLocalMachine.Rows[k]["comments"] = comments;
                                }
                            }
                        }
                    }

                    if (!found || !Properties.Settings.Default.DoNotImportDuplicates)
                    {
                        DataRow dr = frmMain.Instance.dtRemovedLocalMachine.NewRow();

                        dr["icon"] = bmp;
                        dr["name"] = name;
                        dr["setting"] = command;
                        dr["comments"] = comments;

                        frmMain.Instance.dtRemovedLocalMachine.Rows.Add(dr);
                    }
                }
            }
            else if (_ProjectType == ProjectType.HKCU)
            {
                if (dres == DialogResult.No)
                {
                    frmMain.Instance.dtCurrentUser.Clear();
                    frmMain.Instance.dtRemovedCurrentUser.Clear();
                }

                XmlNode no = doc.SelectSingleNode("//CurrentUser");
                XmlNodeList nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string name = nolos[m].Attributes.GetNamedItem("Name").Value;
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    bool found = false;

                    for (int k = 0; k < frmMain.Instance.dtCurrentUser.Rows.Count; k++)
                    {
                        if (frmMain.Instance.dtCurrentUser.Rows[k]["name"].ToString().ToLower() == name.ToLower())
                        {
                            found = true;
                            if (frmMain.Instance.dtCurrentUser.Rows[k]["comments"].ToString() == string.Empty)
                            {
                                frmMain.Instance.dtCurrentUser.Rows[k]["comments"] = comments;
                            }
                        }
                    }

                    if (!found)
                    {
                        DataRow dr = frmMain.Instance.dtCurrentUser.NewRow();

                        Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(command);

                        if (bmp == null)
                        {
                            bmp = EmptyBitmap;
                        }

                        dr["icon"] = bmp;
                        dr["name"] = name;
                        dr["setting"] = command;
                        dr["comments"] = comments;

                        frmMain.Instance.dtCurrentUser.Rows.Add(dr);
                    }
                }                

                no = doc.SelectSingleNode("//RemovedCurrentUser");
                nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string name = nolos[m].Attributes.GetNamedItem("Name").Value;
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    Bitmap bmp = SettingIconExtractor.ExtractIconFromSetting(command);

                    if (bmp == null)
                    {
                        bmp = EmptyBitmap;
                    }
                    bool found = false;

                    for (int k = 0; k < frmMain.Instance.dtRemovedCurrentUser.Rows.Count; k++)
                    {
                        if (frmMain.Instance.dtRemovedCurrentUser.Rows[k]["name"].ToString().ToLower() == name.ToLower())
                        {
                            found = true;
                            if (frmMain.Instance.dtRemovedCurrentUser.Rows[k]["comments"].ToString() == string.Empty)
                            {
                                frmMain.Instance.dtRemovedCurrentUser.Rows[k]["comments"] = comments;
                            }
                        }
                    }

                    if (!found)
                    {
                        DataRow dr = frmMain.Instance.dtRemovedCurrentUser.NewRow();

                        dr["icon"] = bmp;
                        dr["name"] = name;
                        dr["setting"] = command;
                        dr["comments"] = comments;

                        frmMain.Instance.dtRemovedCurrentUser.Rows.Add(dr);
                    }
                }
            }


        }

        public void Save(DataTable dt, DataTable dtRemoved)
        {
            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(_Filepath)))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_Filepath));
            }

            string xml = "<Settings>";

            if (_ProjectType == ProjectType.HKLM)
            {
                xml += "<LocalMachine>";
            }
            else if( _ProjectType==ProjectType.HKCU)
            {
                xml += "<CurrentUser>";
            }

            for (int k = 0; k < dt.Rows.Count; k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(dt.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(dt.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(dt.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            if (_ProjectType == ProjectType.HKLM)
            {
                xml += "</LocalMachine>";
            }
            else if (_ProjectType == ProjectType.HKCU)
            {
                xml += "</CurrentUser>";
            }

            if (_ProjectType == ProjectType.HKLM)
            {
                xml += "<RemovedLocalMachine>";
            }
            else if (_ProjectType == ProjectType.HKCU)
            {
                xml += "<RemovedCurrentUser>";
            }

            for (int k = 0; k < dtRemoved.Rows.Count; k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(dtRemoved.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(dtRemoved.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(dtRemoved.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            if (_ProjectType == ProjectType.HKLM)
            {
                xml += "</RemovedLocalMachine>";
            }
            else if (_ProjectType == ProjectType.HKCU)
            {
                xml += "</RemovedCurrentUser>";
            }            

            xml += "</Settings>";

            System.IO.File.WriteAllText(_Filepath, xml);
        }
    }
}
