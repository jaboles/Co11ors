using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace VSIconSwitcher
{
    public abstract class AssetReplacement
    {
        public AssetReplacement(string filename)
        {
            m_file = filename;
        }

        public virtual void Backup()
        {
            string src = Path.Combine(Options.VS11Folder, FilePath);
            string dest = Path.Combine(Options.BackupFolder, Path.GetFileName(FilePath));

            if (FilePath.Contains('?') || FilePath.Contains('*'))
            {
                IEnumerable<string> matchingFiles = Directory.EnumerateFiles(Path.GetDirectoryName(src), Path.GetFileName(src));
                foreach (string f in matchingFiles)
                {
                    dest = Path.Combine(Options.BackupFolder, Path.GetFileName(f));
                    if (!File.Exists(dest))
                    {
                        File.Copy(f, dest, false);
                    }
                }
            }
            else
            {
                if (!File.Exists(dest))
                {
                    File.Copy(src, dest, false);
                }
            }
        }

        public abstract void DoReplace();

        public virtual void Undo()
        {
            string src = Path.Combine(Options.BackupFolder, Path.GetFileName(FilePath));
            string dest = Path.Combine(Options.VS11Folder, FilePath);

            if (FilePath.Contains('?') || FilePath.Contains('*'))
            {
                IEnumerable<string> matchingFiles = Directory.EnumerateFiles(Path.GetDirectoryName(dest), Path.GetFileName(dest));
                foreach (string f in matchingFiles)
                {
                    src = Path.Combine(Options.BackupFolder, Path.GetFileName(f));
                    if (!File.Exists(f))
                    {
                        File.Copy(src, f, false);
                    }
                }
            }
            else
            {
                File.Copy(src, dest, true);
            }
        }

        public static AssetReplacementOptions Options
        {
            get
            {
                if (s_options == null)
                    s_options = new AssetReplacementOptions();

                return s_options;
            }
        }

        public string FilePath { get { return m_file; } }

        private string m_file;
        private static AssetReplacementOptions s_options;
    }
}
