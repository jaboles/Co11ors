using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Globalization;

namespace VSIconSwitcher
{
    public enum VSAppId
    {
        VisualStudio,
        VWDExpress,
    }

    public class VSInfo
    {
        public VSInfo(string version, VSAppId appId = VSAppId.VisualStudio)
        {
            this.AppID = appId;
            this.Version = version;
            m_installedLanguages = new List<CultureInfo>();
        }

        public static IEnumerable<VSInfo> FindInstalled(string requestedVersion)
        {
            return s_knownVersions.Where(info => info.Version.Equals(requestedVersion) && info.CheckIfExists());
        }

        public VSAppId AppID
        {
            get;
            private set;
        }

        public string Version
        {
            get;
            private set;
        }

        public CultureInfo DefaultLanguage
        {
            get
            {
                if (m_defaultLanguage == null)
                {
                    RegistryKey k = Registry.LocalMachine.OpenSubKey(string.Format(@"Software\Microsoft\{0}\{1}\Setup", AppID, Version));
                    if (k != null)
                    {
                        string firstLanguage = k.GetValue("FirstLanguagePackLCID") as string;
                        k.Close();
                        if (firstLanguage != null)
                        {
                            m_defaultLanguage = CultureInfo.GetCultureInfo(Convert.ToInt32(firstLanguage));
                        }
                        else
                        {
                            throw new Exception("Could not determine the default installed VS language");
                        }
                    }
                    else
                    {
                        throw new Exception("Could not open registry key to the default installed VS language");
                    }
                }
                return m_defaultLanguage;
            }
        }

        public IEnumerable<CultureInfo> InstalledLanguages
        {
            get
            {
                if (m_installedLanguages.Count == 0)
                {
                    RegistryKey k = Registry.LocalMachine.OpenSubKey(string.Format(@"Software\Microsoft\{0}\{1}\Setup\VS\BuildNumber", AppID, Version));
                    if (k != null)
                    {
                        string[] lcids = k.GetValueNames();
                        foreach (string lcid in lcids)
                        {
                            if (!string.IsNullOrWhiteSpace(lcid))
                            {
                                m_installedLanguages.Add(CultureInfo.GetCultureInfo(Convert.ToInt32(lcid)));
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("Could not open registry key to determine installed VS languages");
                    }
                }
                if (m_installedLanguages.Count == 0)
                {
                    throw new Exception("Couldn't find any installed languages, this is a bug");
                }
                return m_installedLanguages;
            }
        }

        public string VSRoot
        {
            get
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(string.Format(@"Software\Microsoft\{0}\{1}", AppID, Version));
                if (k != null)
                {
                    string installDir = k.GetValue("InstallDir") as string;
                    installDir = installDir.Substring(0, installDir.IndexOf("Common7"));
                    k.Close();
                    if (installDir != null)
                    {
                        return installDir;
                    }
                }
                return null;
            }
        }

        public bool CheckIfExists()
        {
            return !string.IsNullOrEmpty(VSRoot);
        }

        private CultureInfo m_defaultLanguage;
        private IList<CultureInfo> m_installedLanguages;

        // A list of the known products
        private static IEnumerable<VSInfo> s_knownVersions = new VSInfo[] {
            new VSInfo("10.0", VSAppId.VisualStudio),
            new VSInfo("11.0", VSAppId.VisualStudio)
        };
    }
}
