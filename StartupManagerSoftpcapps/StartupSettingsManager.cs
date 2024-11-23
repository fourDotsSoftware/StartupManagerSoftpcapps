using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.Data;

namespace StartupManagerSoftpcapps
{
    public class StartupSettingsManager
    {
        public enum StartupSettingsType
        {
            HKLM,
            HKCU
        }

        private StartupSettingsType _StartupSettingsType;
        private DataTable dt = null;

        public StartupSettingsManager(StartupSettingsType startupSettingsType)
        {
            _StartupSettingsType = startupSettingsType;
        }

        public StartupSettingsManager(StartupSettingsType startupSettingsType,DataTable dataTable)
        {
            _StartupSettingsType = startupSettingsType;
            dt = dataTable;
        }

        public string[] GetSettingNames()
        {
            RegistryKey reg = GetRegistryKey(false);

            return reg.GetValueNames();
        }

        public string[] GetSettingValues()
        {
            RegistryKey reg = GetRegistryKey(false);

            string[] s1 = reg.GetValueNames();

            string[] lst = new string[s1.Length];

            for (int k=0;k<s1.Length;k++)
            {
                lst[k] = reg.GetValue(s1[k]).ToString() ;
            }

            return lst;
            
        }

        public RegistryKey GetRegistryKey(bool writeable)
        {
            if (_StartupSettingsType==StartupSettingsType.HKLM)
            {
                return Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run",writeable);
            }
            else if (_StartupSettingsType==StartupSettingsType.HKCU)
            {
                return Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", writeable);
            }
            else
            {
                throw new Exception("Error wrong Startup Setting Type");
            }
            
        }

        public bool ApplySettings()
        {
            RegistryKey reg = null;

            try
            {                
                reg = GetRegistryKey(true);

                string[] vns = reg.GetValueNames();

                for (int k = 0; k < vns.Length; k++)
                {
                    reg.DeleteValue(vns[k]);
                }

                for (int k = 0; k < dt.Rows.Count; k++)
                {
                    reg.SetValue(dt.Rows[k]["name"].ToString(), dt.Rows[k]["setting"].ToString());
                }                

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Error could not Write Settings to Registry !");
            }
            finally
            {
                if (reg!=null)
                {
                    reg.Close();
                }
            }
        }

        public string ValidateSettings()
        {
            for (int k = 0; k < dt.Rows.Count; k++)
            {
                if (dt.Rows[k]["name"].ToString()==string.Empty)
                {
                    return "Error. Empty Name value for Setting !";
                }
                
                if (dt.Rows[k]["setting"].ToString() == string.Empty)
                {
                    return "Error. Empty Setting !";
                }
            }

            return string.Empty;
        }
    }


}
