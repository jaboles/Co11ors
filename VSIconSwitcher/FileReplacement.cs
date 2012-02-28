using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace VSIconSwitcher
{
    public class FileReplacement : AssetReplacement
    {
        public FileReplacement(string filename)
            : base(filename)
        {
        }

        public override void  CopyResources(string sourceFile, string destFile)
        {
            File.Copy(sourceFile, destFile, true);
        }
    }
}
