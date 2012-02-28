using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace VSIconSwitcher
{
    public class AssetReplacementOptions
    {
        public string BackupFolder;
        public string VS10Folder;
        public string VS11Folder;
        public CultureInfo SourceCulture;
        public IEnumerable<CultureInfo> InstalledCultures;
    }
}
