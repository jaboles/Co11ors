using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Mono.Cecil;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Resources;
using System.Collections;

namespace VSIconSwitcher
{
    public class ManagedResourceReplacement : AssetReplacement
    {
        private string m_resourceNamePattern;
        private string m_srcResourcesFilename;
        private string m_dstResourcesFilename;

        public ManagedResourceReplacement(string srcFilename, string dstFilename, string resourcesFilename, string resourceNamePattern)
            : base(srcFilename, dstFilename)
        {
            if (string.IsNullOrWhiteSpace(dstFilename))
            {
                dstFilename = srcFilename;
            }

            if (string.IsNullOrWhiteSpace(resourcesFilename))
            {
                m_srcResourcesFilename = Path.GetFileNameWithoutExtension(srcFilename) + ".g.resources";
                m_dstResourcesFilename = Path.GetFileNameWithoutExtension(dstFilename) + ".g.resources";
            }
            else
            {
                m_srcResourcesFilename = resourcesFilename;
                m_dstResourcesFilename = resourcesFilename;
            }

            // Forward-slash is the correct path separator in resource names, so replace any instances of backslash with it.
            if (resourceNamePattern != null)
            {
                m_resourceNamePattern = resourceNamePattern.Replace('\\', '/');
            }
        }

        public ManagedResourceReplacement(string filename, string resourcesFilename, string resourceNamePattern)
            : this(filename, filename, resourcesFilename, resourceNamePattern)
        {
        }

        public ManagedResourceReplacement(string filename, string resourceNamePattern)
            : this(filename, null, resourceNamePattern)
        {
        }

        public override void CopyResources(string src, string dest)
        {
            AssemblyDefinition srcAD = AssemblyDefinition.ReadAssembly(src);
            AssemblyDefinition dstAD = AssemblyDefinition.ReadAssembly(dest);

            string srcName = srcAD.Name.Name;
            string dstName = dstAD.Name.Name;

            if (string.IsNullOrWhiteSpace(m_resourceNamePattern))
            {
                // Replace the 1st-level resources
                Debug.Assert(string.Equals(m_srcResourcesFilename, m_dstResourcesFilename, StringComparison.OrdinalIgnoreCase));
                Debug.Assert(!string.IsNullOrWhiteSpace(m_srcResourcesFilename));

                m_srcResourcesFilename = m_srcResourcesFilename.Replace("[an]", srcName);

                foreach (EmbeddedResource srcRes in srcAD.MainModule.Resources.OfType<EmbeddedResource>())
                {
                    if (MatchWildcards(m_srcResourcesFilename, srcRes.Name))
                    {
                        string dstResName = srcRes.Name.Replace(srcName, dstName);
                        dstAD.ReplaceResource(dstResName, srcRes.GetResourceData());
                    }
                }
            }
            else
            {
                // Replacing binary resources inside a .resources file

                EmbeddedResource srcBinaryResources;
                EmbeddedResource dstBinaryResources;

                srcBinaryResources = srcAD.MainModule.Resources.OfType<EmbeddedResource>().Single(er => er.Name.Equals(m_srcResourcesFilename, StringComparison.OrdinalIgnoreCase));
                dstBinaryResources = dstAD.MainModule.Resources.OfType<EmbeddedResource>().Single(er => er.Name.Equals(m_dstResourcesFilename, StringComparison.OrdinalIgnoreCase));

                MemoryStream resourceOutputStream = new MemoryStream();
                ResourceReader srcReader = new ResourceReader(srcBinaryResources.GetResourceStream());
                ResourceReader dstReader = new ResourceReader(dstBinaryResources.GetResourceStream());
                ResourceWriter writer = new ResourceWriter(resourceOutputStream);

                IDictionaryEnumerator it = dstReader.GetEnumerator();
                while (it.MoveNext())
                {
                    string currentResourceName = it.Key as string;
                    Debug.Assert(currentResourceName != null);

                    if (MatchWildcards(m_resourceNamePattern, currentResourceName) && srcReader.ContainsResource(currentResourceName))
                    {
                        // Resource name matches wildcard. Copy from patch source
                        UnmanagedMemoryStream data = srcReader.FindResource(currentResourceName);
                        writer.AddResource(currentResourceName, data);
                    }
                    else
                    {
                        // Doesn't match - just copy across.
                        string key = it.Key as string;
                        UnmanagedMemoryStream data = it.Value as UnmanagedMemoryStream;
                        
                        Debug.Assert(key != null);
                        Debug.Assert(data != null);

                        writer.AddResource(key, data);
                    }
                }
                writer.Generate();
                writer.Close();
                dstAD.ReplaceResource(m_dstResourcesFilename, resourceOutputStream.ToArray());
            }
            dstAD.Write(dest);
        }

        public override void BackupItem(string itemPath, string backupLocation)
        {
            base.BackupItem(itemPath, backupLocation);
        }

        public override void UndoItem(string itemPath, string backupLocation)
        {
            base.UndoItem(itemPath, backupLocation);
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

        /// <summary>Convert wildcards to regular expressions and see if it is a match</summary>
        /// <param name="pattern">The filter containing the wildcard(s)</param>
        /// <param name="text">The text we want to match</param>
        /// <returns>true if text matches the pattern</returns>
        private bool MatchWildcards(string pattern, string text)
        {
            pattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
        }
    }
}
