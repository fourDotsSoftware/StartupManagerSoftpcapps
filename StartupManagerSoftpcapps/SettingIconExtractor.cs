using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace StartupManagerSoftpcapps
{
    public class SettingIconExtractor
    {
        public static Bitmap ExtractIconFromSetting(string setting)
        {
            try
            {
                string exefilepath = SettingsParser.GetExeFilepath(setting);

                if (exefilepath == string.Empty) return null;

                ApplicationIconExtractor ae = new ApplicationIconExtractor(exefilepath);

                return ae.Icon;
            }
            catch
            {
                return null;
            }
        }
    }

    public class SettingsParser
    {
        public SettingsParser() { }
        public static string GetExeFilepath(string setting)
        {
            try
            {
                int epos = setting.IndexOf(".exe");

                if (epos < 0)
                {
                    epos = setting.IndexOf(".bat");
                }

                if (epos < 0)
                {
                    return null;
                }

                int qpos = setting.LastIndexOf("\"", epos);

                if (qpos < 0)
                {
                    qpos = setting.LastIndexOf("'", epos);

                    if (qpos < 0)
                    {
                        qpos = 0;
                    }
                    else
                    {
                        qpos++;
                    }
                }
                else
                {
                    qpos++;
                }

                //01234
                string exefilepath = setting.Substring(qpos, epos + 4 - qpos);

                return exefilepath;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
