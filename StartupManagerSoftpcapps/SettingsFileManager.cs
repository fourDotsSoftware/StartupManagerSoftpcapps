using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Security;
using System.Data;
using System.Drawing;

namespace StartupManagerSoftpcapps
{
    public class SettingsFileManager    
    {
        public string SettingsFile =
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + Module.ApplicationName + "\\settings.xml";

        private Bitmap EmptyBitmap = new Bitmap(16, 16);

        public void SaveSettings()
        {
            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(SettingsFile)))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(SettingsFile));
            }

            string xml = "<Settings>";

            xml += "<LocalMachine>";

            for (int k=0;k<frmMain.Instance.dtLocalMachine.Rows.Count;k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(frmMain.Instance.dtLocalMachine.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(frmMain.Instance.dtLocalMachine.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(frmMain.Instance.dtLocalMachine.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            xml += "</LocalMachine>";

            xml += "<RemovedLocalMachine>";

            for (int k = 0; k < frmMain.Instance.dtRemovedLocalMachine.Rows.Count; k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedLocalMachine.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedLocalMachine.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedLocalMachine.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            xml += "</RemovedLocalMachine>";

            xml += "<CurrentUser>";

            for (int k = 0; k < frmMain.Instance.dtCurrentUser.Rows.Count; k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(frmMain.Instance.dtCurrentUser.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(frmMain.Instance.dtCurrentUser.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(frmMain.Instance.dtCurrentUser.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            xml += "</CurrentUser>";

            xml += "<RemovedCurrentUser>";

            for (int k = 0; k < frmMain.Instance.dtRemovedCurrentUser.Rows.Count; k++)
            {
                xml += "<Setting Name=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedCurrentUser.Rows[k]["name"].ToString())
                    + "\" Command=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedCurrentUser.Rows[k]["setting"].ToString())
                    + "\" Comments=\"" + SecurityElement.Escape(frmMain.Instance.dtRemovedCurrentUser.Rows[k]["comments"].ToString())
                    + "\" />";
            }

            xml += "</RemovedCurrentUser>";

            xml+= "</Settings>";

            System.IO.File.WriteAllText(SettingsFile, xml);
        }

        public void LoadSettings()
        {
            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(SettingsFile)))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(SettingsFile));
            }

            if (System.IO.File.Exists(SettingsFile))
            {
                XmlDocument doc = new XmlDocument();

                doc.LoadXml(System.IO.File.ReadAllText(SettingsFile));

                XmlNode no = doc.SelectSingleNode("//LocalMachine");
                XmlNodeList nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    for (int k = 0; k < frmMain.Instance.dtLocalMachine.Rows.Count; k++)
                    {
                        if (frmMain.Instance.dtLocalMachine.Rows[k]["setting"].ToString().ToLower() == command.ToLower())
                        {
                            if (frmMain.Instance.dtLocalMachine.Rows[k]["comments"].ToString() == string.Empty)
                            {
                                frmMain.Instance.dtLocalMachine.Rows[k]["comments"] = comments;
                            }
                        }
                    }
                }

                no = doc.SelectSingleNode("//CurrentUser");
                nolos = no.SelectNodes("./Setting");

                for (int m = 0; m < nolos.Count; m++)
                {
                    string command = nolos[m].Attributes.GetNamedItem("Command").Value;
                    string comments = nolos[m].Attributes.GetNamedItem("Comments").Value;

                    for (int k = 0; k < frmMain.Instance.dtCurrentUser.Rows.Count; k++)
                    {
                        if (frmMain.Instance.dtCurrentUser.Rows[k]["setting"].ToString().ToLower() == command.ToLower())
                        {
                            if (frmMain.Instance.dtCurrentUser.Rows[k]["comments"].ToString() == string.Empty)
                            {
                                frmMain.Instance.dtCurrentUser.Rows[k]["comments"] = comments;
                            }
                        }
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

                    DataRow dr = frmMain.Instance.dtRemovedLocalMachine.NewRow();

                    dr["icon"] = bmp;
                    dr["name"] = name;
                    dr["setting"] = command;
                    dr["comments"] = comments;

                    frmMain.Instance.dtRemovedLocalMachine.Rows.Add(dr);
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

    
}
