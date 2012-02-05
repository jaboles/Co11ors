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

        public override void DoReplace()
        {
            string src = Path.Combine(Options.VS10Folder, FilePath);
            string dest = Path.Combine(Options.VS11Folder, FilePath);

            IEnumerable<string> srcFiles = Directory.EnumerateFiles(Path.GetDirectoryName(src), Path.GetFileName(src));
            IEnumerable<string> dstFiles = Directory.EnumerateFiles(Path.GetDirectoryName(dest), Path.GetFileName(dest));


            string srcFilesStr = string.Join(Environment.NewLine, srcFiles);
            string dstFilesStr = string.Join(Environment.NewLine, dstFiles);
            Debug.Assert(srcFiles.SequenceEqual(dstFiles), "Set of source files doesn't match destination files");

            foreach (string s in srcFiles)
            {

            }
        }
    }
}
