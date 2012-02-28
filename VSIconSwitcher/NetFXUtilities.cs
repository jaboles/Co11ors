using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace VSIconSwitcher
{
    public class NetFXUtilities
    {
        public static bool IsInGAC(string assemblyName)
        {
            string output = ExecGacUtil(string.Format("-l {0}", GetAssemblyName(assemblyName)));
            if (output.Contains("Number of items = 0"))
            {
                return false;
            }
            else if (output.Contains("Number of items = "))
            {
                return true;
            }
            throw new Exception("Unknown output received from gacutil.exe: " + output);
        }

        public static void SkipVerification(string assemblyName)
        {
            string output = ExecGacUtil(string.Format("-Vr {0}", GetAssemblyName(assemblyName)));
            if (output.Contains("Verification entry added for assembly"))
            {
                return;
            }
            throw new Exception("Unknown output received from gacutil.exe: " + output);
        }

        public static void UnskipVerification(string assemblyName)
        {
            string output = ExecGacUtil(string.Format("-Vu {0}", GetAssemblyName(assemblyName)));
            if (output.Contains("Verification entry for assembly"))
            {
                return;
            }
            throw new Exception("Unknown output received from gacutil.exe: " + output);
        }

        private static string ExecGacUtil(string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = GacUtilExe;
            psi.Arguments = arguments;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.UseShellExecute = false;
            Process p = Process.Start(psi);
            p.WaitForExit();
            if (p.ExitCode != 0)
            {
                throw new Exception("Gacutil.exe returned an error: " + p.ExitCode.ToString());
            }

            string output = p.StandardOutput.ReadToEnd();
            return output;
        }

        public static string GacUtilExe
        {
            get
            {
                if (s_gacUtilExe == null)
                {
                    string[] gacUtils = new string[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft SDKs", "Windows", "v7.0A", "Bin", "NETFX 4.0 Tools", "gacutil.exe"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft SDKs", "Windows", "v7.0A", "Bin", "gacutil.exe"),
                    };

                    foreach (string s in gacUtils)
                    {
                        if (File.Exists(s))
                        {
                            s_gacUtilExe = s;
                            break;
                        }
                    }
                }
                return s_gacUtilExe;
            }
        }

        private static string GetAssemblyName(string assemblyName)
        {
            if (assemblyName.Contains(Path.DirectorySeparatorChar))
                assemblyName = Path.GetFileName(assemblyName);

            if (assemblyName.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
                assemblyName = assemblyName.Substring(0, assemblyName.Length - 4);

            return assemblyName;
        }

        private static string s_gacUtilExe;
    }
}
