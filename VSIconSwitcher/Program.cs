using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Globalization;

namespace VSIconSwitcher
{
    public class Program
    {
        private MainForm form;
        private VSInfo m_srcVS;
        private VSInfo m_dstVS;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            new Program();
        }

        public Program()
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

            m_srcVS = VSInfo.FindInstalled("10.0").First();
            m_dstVS = VSInfo.FindInstalled("11.0").First();

            form.VS10Path = m_srcVS.VSRoot;
            form.VS11Path = m_dstVS.VSRoot;
            form.BackupPath = DefaultBackupPath;

            foreach (CultureInfo ci in m_srcVS.InstalledLanguages)
                System.Windows.Forms.MessageBox.Show(ci.LCID.ToString());

            Application.Run(form);
        }

        void BackupAndPatch(string vs10Path, string vs11Path, string backupPath)
        {
            var context = TaskScheduler.FromCurrentSynchronizationContext();

            AssetReplacement.Options.BackupFolder = backupPath;
            AssetReplacement.Options.VS10Folder = vs10Path;
            AssetReplacement.Options.VS11Folder = vs11Path;

            form.ProgressMax = 2 * ReplacementList.VS11Ultimate.Count();
            form.CurrentProgress = 0;

            form.IsBusy = true;
            form.Status = "Backing up files";
            if (!Directory.Exists(backupPath))
            {
                Directory.CreateDirectory(backupPath);
            }
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                Task.Factory.StartNew(() =>
                {
                    r.Backup();
                }, CancellationToken.None, TaskCreationOptions.None, Scheduler).ContinueWith((task) =>
                {
                    form.CurrentProgress++;
                }, context);
            }
            form.Status = "Patching resources";
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                Task.Factory.StartNew(() =>
                {
                    r.DoReplace();
                }, CancellationToken.None, TaskCreationOptions.None, Scheduler).ContinueWith((task) =>
                {
                    form.CurrentProgress++;
                }, context);
            }
            form.Status = "Running devenv /setup";
            Task.Factory.StartNew(() =>
            {
                RunDevenvSetup(vs11Path);
            }, CancellationToken.None, TaskCreationOptions.None, Scheduler).ContinueWith((task) =>
            {
                form.Status = null;
                form.CurrentProgress = 0;
                form.IsBusy = false;
            }, context);
        }

        void Undo(string vs11Path, string backupPath)
        {
            var context = TaskScheduler.FromCurrentSynchronizationContext();

            AssetReplacement.Options.BackupFolder = backupPath;
            AssetReplacement.Options.VS11Folder = vs11Path;

            form.Status = "Undoing changes";
            form.CurrentProgress = 0;
            form.ProgressMax = ReplacementList.VS11Ultimate.Count();

            form.IsBusy = true;
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                Task.Factory.StartNew(() =>
                {
                    r.Undo();
                }, CancellationToken.None, TaskCreationOptions.None, Scheduler).ContinueWith((task) =>
                {
                    form.CurrentProgress++;
                }, context);
            }

            form.Status = "Running devenv /setup";
            Task.Factory.StartNew(() =>
            {
                RunDevenvSetup(vs11Path);
            }, CancellationToken.None, TaskCreationOptions.None, Scheduler).ContinueWith((task) =>
            {
                form.Status = null;
                form.CurrentProgress = 0;
                form.IsBusy = false;
            }, context);
        }

        string DefaultVS10Path
        {
            get
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\10.0");
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

        string DefaultVS11Path
        {
            get
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\11.0");
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

        void RunDevenvSetup(string folder)
        {
            Process p = Process.Start(Path.Combine(folder, "Common7\\IDE\\devenv.exe"), "/setup");
            p.WaitForExit();
        }

        string DefaultBackupPath
        {
            get
            {
                return Path.Combine(Path.GetTempPath(), "VsIconSwitcherBackup");
            }
        }

        TaskScheduler Scheduler
        {
            get
            {
                if (m_scheduler == null)
                {
                    m_scheduler = new LimitedConcurrencyLevelTaskScheduler(1);
                }
                return m_scheduler;
            }
        }

        private TaskScheduler m_scheduler;
    }
}
