﻿using System;
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
            IEnumerable files = Directory.EnumerateFiles("c:\\Program Files (x86)\\Microsoft Visual Studio 11.0 - Copy - Copy", "*.dll", SearchOption.AllDirectories);
            foreach (string f in files)
            {
                Console.WriteLine("Checking assembly: {0}", f);
                if (IsManaged(f))
                {
                    Console.WriteLine("Found managed assembly. Processing...", f);
                    new Program(f);
                }
            }
        }

        public Program(string path)
        {
            AssemblyDefinition ad = AssemblyDefinition.ReadAssembly(path);
            Assembly dotNetAssembly = Assembly.LoadFrom(path);
            string[] asdf = dotNetAssembly.GetManifestResourceNames();
            IList<string> unprocessedAssemblies = new List<string>();
            foreach (Resource res in ad.MainModule.Resources)
            {
                EmbeddedResource er = res as EmbeddedResource;
                Console.WriteLine("Resource name: {0}", er.Name);

                string outputFolder = Path.Combine("c:\\mrd", Path.GetFileName(path), dotNetAssembly.GetName().Version.ToString());
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                if (res.Name.ToLower().EndsWith(".resources"))
                {
                    string resourceFileBaseName = res.Name.Substring(0, res.Name.Length - 10);

                    ResourceManager rm = new ResourceManager(resourceFileBaseName, dotNetAssembly);
                    try
                    {
                        ResourceSet rs = rm.GetResourceSet(CultureInfo.CurrentCulture, true, true);
                        IDictionaryEnumerator re = rs.GetEnumerator();
                        while (re.MoveNext())
                        {
                            Stream resStream = re.Value as Stream;
                            if (resStream != null)
                            {
                                string outputDir = Path.Combine(outputFolder, resourceFileBaseName, Path.GetDirectoryName(re.Key.ToString()));
                                if (!Directory.Exists(outputDir))
                                    Directory.CreateDirectory(outputDir);

                                string outputPath = Path.Combine(outputDir, Path.GetFileName(re.Key.ToString()));

                                byte[] data = new byte[resStream.Length];
                                resStream.Read(data, 0, (int)resStream.Length);

                                FileStream f = new FileStream(outputPath, FileMode.Create);
                                f.Write(data, 0, data.Length);
                                f.Close();
                            }
                        }
                    }
                    catch (MissingManifestResourceException)
                    {
                        unprocessedAssemblies.Add(res.Name);
                    }
                    catch (MissingSatelliteAssemblyException)
                    {
                        unprocessedAssemblies.Add(res.Name);
                    }
                }
                else
                {
                    byte[] data = er.GetResourceData();

                    string outputPath = Path.Combine(outputFolder, Sanitize(res.Name));

                    FileStream f = new FileStream(outputPath, FileMode.Create);
                    f.Write(data, 0, data.Length);
                    f.Close();
                }
            }
            if (unprocessedAssemblies.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("UNPROCESSED RESOURCES:");
                foreach (string s in unprocessedAssemblies)
                {
                    Console.WriteLine("  {0}", s);
                }
            }
        }

        public static bool IsManaged(string assemblyPath)
        {
            try
            {
                AssemblyName.GetAssemblyName(assemblyPath);
                return true;
            }
            catch (BadImageFormatException)
            {
                return false;
            }
        }

        public string Sanitize(string path)
        {
            return path.Replace(':', '_');
        }
    }
}
