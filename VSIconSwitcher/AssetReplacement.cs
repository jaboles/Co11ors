using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Globalization;

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
            if (dstFilename == null)
                dstFilename = srcFilename;
            m_dstFile = dstFilename;
        }

        /// <summary>
        /// Expands filenames
        /// </summary>
        /// <param name="filenamePattern"></param>
        /// <returns>Returns an enumerable of string-pairs which contain the expanded filename, and the backup location.</returns>
        public IEnumerable<Tuple<string, string>> ExpandFilenames(string directory, string filenamePattern)
        {
            List<Tuple<string, string>> list = new List<Tuple<string,string>>();

            // Need to do culture token expansion first, so that Directory.EnumerateFiles can be used to do wildcard expansion.
            IEnumerable<Tuple<string, string>> localisedFilenamePatterns = ReplaceCultureTokens(filenamePattern);

            foreach (Tuple<string, string> culturisedItem in localisedFilenamePatterns)
            {
                string localisedFilenamePattern = culturisedItem.Item1;
                string cultureSpecificBackupFolder = culturisedItem.Item2;

                if (filenamePattern.Contains('?') || filenamePattern.Contains('*'))
                {
                    string fullPattern = Path.Combine(directory, localisedFilenamePattern);
                    IEnumerable<string> matchingFiles = Directory.EnumerateFiles(Path.GetDirectoryName(fullPattern), Path.GetFileName(fullPattern));
                    foreach (string f in matchingFiles)
                    {
                        string backupLocation = Path.Combine(Options.BackupFolder, Path.GetFileName(f));
                        list.Add(new Tuple<string, string>(f, backupLocation));
                    }
                }
                else
                {
                    string backupLocation = Path.Combine(Options.BackupFolder, Path.GetFileName(localisedFilenamePattern));
                    string f = Path.Combine(directory, localisedFilenamePattern);
                    list.Add(new Tuple<string, string>(f, backupLocation));
                }
            }
            return list;
        }

        public virtual void Backup()
        {
            foreach (Tuple<string, string> fileCopyItem in ExpandFilenames(Options.VS11Folder, DestFilePath))
            {
                string file = fileCopyItem.Item1;
                string backupLocation = fileCopyItem.Item2;

                BackupItem(file, backupLocation);
            }
        }

        public virtual void BackupItem(string itemPath, string backupLocation)
        {
            if (!File.Exists(backupLocation))
            {
                FileCopyAndPreserveDate(itemPath, backupLocation, false);
            }
        }

        public void DoReplace()
        {
            foreach (Tuple<string, string> item in ExpandFilenames(Options.VS11Folder, DestFilePath))
            {
                string resSrc = ExpandFilenames(Options.VS10Folder, SourceFilePath).First().Item1;
                string resDst = item.Item1;

                CopyResources(resSrc, resDst);
            }
        }

        public abstract void CopyResources(string sourceFile, string destFile);

        public virtual void Undo()
        {
            foreach (Tuple<string, string> fileCopyItem in ExpandFilenames(Options.VS11Folder, DestFilePath))
            {
                string file = fileCopyItem.Item1;
                string backupLocation = fileCopyItem.Item2;

                UndoItem(file, backupLocation);
            }
        }

        public virtual void UndoItem(string itemPath, string backupLocation)
        {
            if (File.Exists(backupLocation))
            {
                FileCopyAndPreserveDate(backupLocation, itemPath, true);
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

        protected void FileCopyAndPreserveDate(string src, string dst, bool overwrite)
        {
            File.Copy(src, dst, overwrite);
            FileInfo srcInfo = new FileInfo(src);
            FileInfo dstInfo = new FileInfo(dst);
            dstInfo.CreationTimeUtc = srcInfo.CreationTimeUtc;
            dstInfo.LastWriteTimeUtc = srcInfo.LastWriteTimeUtc;
        }

        private IEnumerable<Tuple<string, string>> ReplaceCultureTokens(string path)
        {
            IList<Tuple<string, string>> list = new List<Tuple<string, string>>();
            if (ContainsCultureTokens(path))
            {
                foreach (CultureInfo c in Options.InstalledCultures)
                {
                    list.Add(ReplaceCultureTokens(path, c));
                }
            }
            else
            {
                // No replacement necessary
                list.Add(new Tuple<string,string>(path, string.Empty));
            }
            return list;
        }

        private Tuple<string, string> ReplaceCultureTokens(string path, CultureInfo cultureInfo)
        {
            path = path.Replace("[lcid]", cultureInfo.LCID.ToString());
            return new Tuple<string, string>(path, cultureInfo.LCID.ToString());
        }

        public bool ContainsCultureTokens(string path)
        {
            return path.Contains("[lcid]");
        }

        public string FilePath { get { Debug.Assert(m_srcFile.Equals(m_dstFile), "Source file and destination file must match if using the FilePath property"); return m_srcFile; } }
        public string SourceFilePath { get { return m_srcFile; } }
        public string DestFilePath { get { return m_dstFile; } }

        private string m_srcFile;
        private string m_dstFile;
        private static AssetReplacementOptions s_options;
    }
}
