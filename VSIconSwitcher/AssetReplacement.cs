using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace VSIconSwitcher
{
    public abstract class AssetReplacement
    {
        public AssetReplacement(string filename)
        {
            m_srcFile = filename;
            m_dstFile = filename;
        }

        public AssetReplacement(string srcFilename, string dstFilename)
        {
            m_srcFile = srcFilename;
            m_dstFile = srcFilename;
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

        public string FilePath { get { Debug.Assert(m_srcFile.Equals(m_dstFile), "Source file and destination file must match if using the FilePath property"); return m_srcFile; } }
        public string SourceFilePath { get { return m_srcFile; } }
        public string DestFilePath { get { return m_dstFile; } }

        private string m_srcFile;
        private string m_dstFile;
        private static AssetReplacementOptions s_options;
    }
}
