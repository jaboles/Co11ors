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
            foreach (Resource res in ad.MainModule.Resources)
            {
                EmbeddedResource er = res as EmbeddedResource;
                Console.WriteLine("Resource name: {0}", er.Name);
                byte[] data = er.GetResourceData();

                string outputFolder = Path.Combine(Path.GetTempPath(), "mrd", Path.GetFileName(path));
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string resourceOutputFolder = Path.Combine(outputFolder, res.Name);
                if (!Directory.Exists(resourceOutputFolder))
                    Directory.CreateDirectory(resourceOutputFolder);

                ResourceManager rm = new ResourceManager(res.Name, dotNetAssembly);
                ResourceSet rs = rm.GetResourceSet(CultureInfo.CurrentCulture, false, false);
                IDictionaryEnumerator re = rs.GetEnumerator();
                while (re.MoveNext())
                {
                    Console.WriteLine("Key: {0}", re.Key);
                    Console.WriteLine("Value: {0}", re.Value);
                    Console.WriteLine();
                }

                /*string outputPath = Path.Combine(outputFolder, res.Name);
                FileStream f = new FileStream(outputPath, FileMode.Create);
                f.Write(data, 0, data.Length);
                f.Close();*/
            }
        }
    }
}
