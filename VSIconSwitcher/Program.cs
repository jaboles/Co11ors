using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VSIconSwitcher
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            MainForm form = new MainForm();
            form.PatchButtonClicked += (o, e) =>
            {
                BackupAndPatch(form.VS10Path, form.VS11Path, form.BackupPath);
            };

            form.UndoButtonClicked += (o, e) =>
            {
                Undo(form.VS10Path, form.BackupPath);
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
        }

        static void Undo(string vs10Path, string backupPath)
        {
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

        static string DefaultBackupPath
        {
            get
            {
                return Path.Combine(Path.GetTempPath(), "VsIconSwitcherBackup");
            }
        }
    }
}
