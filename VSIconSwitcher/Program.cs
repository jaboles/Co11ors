﻿using System;
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

            m_srcVS = new VSInfo("10.0");
            m_dstVS = new VSInfo("11.0");

            form.VS10Path = m_srcVS.VSRoot;
            form.VS11Path = m_dstVS.VSRoot;
            form.BackupPath = DefaultBackupPath;

            foreach (CultureInfo ci in m_srcVS.InstalledLanguages)
                System.Windows.Forms.MessageBox.Show(ci.LCID.ToString());

            Application.Run(form);
        }

        void BackupAndPatch(string vs10Path, string vs11Path, string backupPath)
        {
            ResourceReplacement.Options.BackupFolder = backupPath;
            ResourceReplacement.Options.VS10Folder = vs10Path;
            ResourceReplacement.Options.VS11Folder = vs11Path;

            form.ProgressMax = 2 * ReplacementList.VS11Ultimate.Count();
            form.CurrentProgress = 0;

            try
            {
                form.Enabled = false;
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
            finally
            {
                form.Enabled = true;
            }
        }

        void Undo(string vs11Path, string backupPath)
        {
            ResourceReplacement.Options.BackupFolder = backupPath;
            ResourceReplacement.Options.VS11Folder = vs11Path;

            form.Status = "Undoing changes";
            form.CurrentProgress = 0;
            form.ProgressMax = ReplacementList.VS11Ultimate.Count();

            try
            {
                form.Enabled = false;
                foreach (ResourceReplacement r in ReplacementList.VS11Ultimate)
                {
                    r.Undo();
                    form.CurrentProgress++;
                }

                form.Status = "Running devenv /setup";
                RunDevenvSetup(vs11Path);
            }
            finally
            {
                form.Enabled = true;
            }
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
            var syncContext = TaskScheduler.FromCurrentSynchronizationContext();

            Task.Factory.StartNew(() =>
            {
                Process p = Process.Start(Path.Combine(folder, "Common7\\IDE\\devenv.exe"), "/setup");
                p.WaitForExit();
            }).ContinueWith(t =>
            {
                form.Status = null;
                form.CurrentProgress = 0;
            }, syncContext);
        }

        string DefaultBackupPath
        {
            get
            {
                return Path.Combine(Path.GetTempPath(), "VsIconSwitcherBackup");
            }
        }
    }
}
