using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VSIconSwitcher
{
    static class Program
    {
        static MainForm form;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            form = new MainForm();
            form.PatchButtonClicked += (o, e) =>
            {
                BackupAndPatch(form.VS10Path, form.VS11Path, form.BackupPath);
            };

            form.UndoButtonClicked += (o, e) =>
            {
                Undo(form.VS11Path, form.BackupPath);
            };

            form.QuitButtonClicked += (o, e) =>
            {
                Application.Exit();
            };

            form.VS10Path = DefaultVS10Path;
            form.VS11Path = DefaultVS11Path;
            form.BackupPath = DefaultBackupPath;

            Application.Run(form);
        }

        static void BackupAndPatch(string vs10Path, string vs11Path, string backupPath)
        {
            ResourceReplacement.Options.BackupFolder = backupPath;
            ResourceReplacement.Options.VS10Folder = vs10Path;
            ResourceReplacement.Options.VS11Folder = vs11Path;

            form.ProgressMax = 2 * ReplacementList.VS11Ultimate.Length;
            form.CurrentProgress = 0;

            if (!Directory.Exists(backupPath))
            {
                Directory.CreateDirectory(backupPath);
            }

            form.Status = "Backing up files";
            foreach (ResourceReplacement r in ReplacementList.VS11Ultimate)
            {
                r.Backup();
                form.CurrentProgress++;
            }
            form.Status = "Patching resources";
            foreach (ResourceReplacement r in ReplacementList.VS11Ultimate)
            {
                r.DoReplace();
                form.CurrentProgress++;
            }
            form.Status = "Running devenv /setup";
            RunDevenvSetup(vs11Path);
        }

        static void Undo(string vs11Path, string backupPath)
        {
            ResourceReplacement.Options.BackupFolder = backupPath;
            ResourceReplacement.Options.VS11Folder = vs11Path;

            form.Status = "Undoing changes";
            form.CurrentProgress = 0;
            form.ProgressMax = ReplacementList.VS11Ultimate.Length;

            foreach (ResourceReplacement r in ReplacementList.VS11Ultimate)
            {
                r.Undo();
                form.CurrentProgress++;
            }

            form.Status = "Running devenv /setup";
            RunDevenvSetup(vs11Path);
        }

        static string DefaultVS10Path
        {
            get
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\10.0");
                if (k != null)
                {
                    string installDir = k.GetValue("InstallDir") as string;
                    k.Close();
                    if (installDir != null)
                    {
                        return installDir;
                    }
                }
                return null;
            }
        }

        static string DefaultVS11Path
        {
            get
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\11.0");
                if (k != null)
                {
                    string installDir = k.GetValue("InstallDir") as string;
                    k.Close();
                    if (installDir != null)
                    {
                        return installDir;
                    }
                }
                return null;
            }
        }

        static void RunDevenvSetup(string folder)
        {
            var syncContext = TaskScheduler.FromCurrentSynchronizationContext();

            Task.Factory.StartNew(() =>
            {
                Process p = Process.Start(Path.Combine(folder, "devenv.exe"), "/setup");
                p.WaitForExit();
            }).ContinueWith(t =>
            {
                form.Status = null;
            }, syncContext);
        }

        static string DefaultBackupPath
        {
            get
            {
                return Path.Combine(Path.GetTempPath(), "VsIconSwitcherBackup");
            }
        }
    }
}
