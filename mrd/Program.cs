using System;
using System.Collections.Generic;
using System.Reflection;
using System.Resources;
using System.Linq;
using System.Text;
using Mono.Cecil;
using System.IO;
using System.Globalization;
using System.Collections;

namespace mrd
{
    class Program
    {
        static void Main(string[] args)
        {
            string folder = @"C:\Program Files (x86)\Microsoft Visual Studio 11.0 - Copy - Copy\Common7\IDE\Extensions\Microsoft\TeamFoundation\Team Explorer";
            string path = Path.Combine(folder, "Microsoft.TeamFoundation.CodeReview.Components.dll");

            AssemblyDefinition ad = AssemblyDefinition.ReadAssembly(path);
            Assembly dotNetAssembly = Assembly.LoadFrom(path);
            string[] asdf = dotNetAssembly.GetManifestResourceNames();
            foreach (Resource res in ad.MainModule.Resources)
            {
                EmbeddedResource er = res as EmbeddedResource;
                Console.WriteLine("Resource name: {0}", er.Name);

                string outputFolder = Path.Combine(Path.GetTempPath(), "mrd", Path.GetFileName(path));
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string resourceFileBaseName = res.Name.Substring(0, res.Name.Length - 10);

                ResourceManager rm = new ResourceManager(resourceFileBaseName, dotNetAssembly);
                ResourceSet rs = rm.GetResourceSet(CultureInfo.CurrentCulture, true, true);
                IDictionaryEnumerator re = rs.GetEnumerator();
                while (re.MoveNext())
                {
                    Console.WriteLine("Key: {0}", re.Key);
                    Console.WriteLine("Value: {0}", re.Value);
                    Console.WriteLine();

                    string outputDir = Path.Combine(outputFolder, resourceFileBaseName, Path.GetDirectoryName(re.Key.ToString()));
                    if (!Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    string outputPath = Path.Combine(outputDir, Path.GetFileName(re.Key.ToString()));
                    Stream resStream = re.Value as Stream;
                    if (resStream != null)
                    {
                        byte[] data = new byte[resStream.Length];
                        resStream.Read(data, 0, (int)resStream.Length);

                        FileStream f = new FileStream(outputPath, FileMode.Create);
                        f.Write(data, 0, data.Length);
                        f.Close();
                    }
                }

            }
        }
    }
}
