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
using System.ComponentModel;

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
                BackupAndPatch();
            };

            form.UndoButtonClicked += (o, e) =>
            {
                Undo();
            };

            form.QuitButtonClicked += (o, e) =>
            {
                Application.Exit();
            };

            m_srcVS = VSInfo.FindInstalled("10.0").First();
            m_dstVS = VSInfo.FindInstalled("11.0").First();

            form.VS10Path = m_srcVS.VSRoot;
            form.VS10DetectedProduct = m_srcVS.Version;
            form.VS10Languages = string.Join(", ", m_srcVS.InstalledLanguages.Select(c => c.Name));
            form.VS11Path = m_dstVS.VSRoot;
            form.VS11DetectedProduct = m_dstVS.Version;
            form.VS11Languages = string.Join(", ", m_dstVS.InstalledLanguages.Select(c => c.Name));
            form.BackupPath = DefaultBackupPath;

            Application.Run(form);
        }

        void BackupAndPatch()
        {
            SetReplacementOptions();

            if (!KillDevenvIfRunning()) return;
            
            form.ProgressMax = 2 * ReplacementList.VS11Ultimate.Count();
            form.CurrentProgress = 0;

            form.IsBusy = true;
            form.Status = "Backing up files";
            if (!Directory.Exists(AssetReplacement.Options.BackupFolder))
            {
                Directory.CreateDirectory(AssetReplacement.Options.BackupFolder);
            }
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                form.Status = string.Format("Backing up {0}", r.DestFilePath);
                r.Backup();
                form.CurrentProgress++;
            }
            form.Status = "Patching resources";
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                form.Status = string.Format("Patching resources in {0}", r.DestFilePath);
                r.DoReplace();
                form.CurrentProgress++;
            }
            form.Status = "Running devenv /setup";
            RunDevenvSetup(AssetReplacement.Options.VS11Folder);

            form.Status = null;
            form.CurrentProgress = 0;
            form.IsBusy = false;
        }

        void Undo()
        {
            SetReplacementOptions();

            if (!KillDevenvIfRunning()) return;

            form.Status = "Undoing changes";
            form.CurrentProgress = 0;
            form.ProgressMax = ReplacementList.VS11Ultimate.Count();

            form.IsBusy = true;
            foreach (AssetReplacement r in ReplacementList.VS11Ultimate)
            {
                form.Status = string.Format("Undoing changes to {0}", r.DestFilePath);
                r.Undo();
                form.CurrentProgress++;
            }

            form.Status = "Running devenv /setup";
            RunDevenvSetup(AssetReplacement.Options.VS11Folder);

            form.Status = null;
            form.CurrentProgress = 0;
            form.IsBusy = false;
        }

        void SetReplacementOptions()
        {
            AssetReplacement.Options.BackupFolder = form.BackupPath;
            AssetReplacement.Options.VS10Folder = form.VS10Path;
            AssetReplacement.Options.VS11Folder = form.VS11Path;
            AssetReplacement.Options.SourceCulture = m_srcVS.DefaultLanguage;
            AssetReplacement.Options.InstalledCultures = m_dstVS.InstalledLanguages;
        }

        void RunDevenvSetup(string folder)
        {
            Process p = Process.Start(m_dstVS.VSExe, "/setup");
            p.WaitForExit();
        }

        bool KillDevenvIfRunning()
        {
            Func<IEnumerable<Process>> GetDstVSRunning = () => Process.GetProcesses().Where(proc =>
            {
                if (!m_dstVS.VSExe.ToLower().Contains(proc.ProcessName.ToLower())) return false;
                try
                {
                    return proc.MainModule.FileName.Equals(m_dstVS.VSExe, StringComparison.OrdinalIgnoreCase);
                }
                catch (Win32Exception)
                {
                    // Probably a 64-bit process.
                    return false;
                }
            });

            if (GetDstVSRunning().Count() > 0)
            {
                DialogResult dr = MessageBox.Show("At least one instance of Visual Studio 11 is running. Kill?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    foreach (Process p in GetDstVSRunning())
                    {
                        p.Kill();
                    }
                    SpinWait.SpinUntil(() => GetDstVSRunning().Count() == 0, TimeSpan.FromSeconds(10));
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
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
