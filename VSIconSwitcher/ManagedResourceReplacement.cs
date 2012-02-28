using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Mono.Cecil;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace VSIconSwitcher
{
    public class ManagedResourceReplacement : AssetReplacement
    {
        private string m_resourceNamePattern;
        private string m_resourcesFilename;

        public ManagedResourceReplacement(string srcFilename, string dstFilename, string resourcesFilename, string resourceNamePattern)
            : base(srcFilename, dstFilename)
        {
            m_resourcesFilename = resourcesFilename;
            m_resourceNamePattern = resourceNamePattern;
        }

        public ManagedResourceReplacement(string filename, string resourcesFilename, string resourceNamePattern)
            : base(filename)
        {
            m_resourcesFilename = resourcesFilename;
            m_resourceNamePattern = resourceNamePattern;
        }

        public ManagedResourceReplacement(string filename, string resourceNamePattern)
            : base(filename)
        {
            m_resourcesFilename = Path.GetFileNameWithoutExtension(filename) + ".g.resources";
            m_resourceNamePattern = resourceNamePattern;
        }

        public override void CopyResources(string src, string dest)
        {
            AssemblyDefinition srcAD = AssemblyDefinition.ReadAssembly(src);
            AssemblyDefinition dstAD = AssemblyDefinition.ReadAssembly(dest);

            if (string.IsNullOrWhiteSpace(m_resourceNamePattern))
            {
                // Replace the 1st-level resources
                Debug.Assert(!string.IsNullOrWhiteSpace(m_resourcesFilename));

                foreach (EmbeddedResource srcRes in srcAD.MainModule.Resources.OfType<EmbeddedResource>())
                {
                    string resName = srcRes.Name;
                    if (MatchWildcards(m_resourcesFilename, resName))
                    {
                        byte[] data = srcRes.GetResourceData();
                        Resource resourceToReplace = dstAD.MainModule.Resources.Single(r => r.Name.Equals(resName));
                        ManifestResourceAttributes attribs = resourceToReplace.Attributes;
                        int index = dstAD.MainModule.Resources.IndexOf(resourceToReplace);

                        dstAD.MainModule.Resources.RemoveAt(index);
                        dstAD.MainModule.Resources.Insert(index, new EmbeddedResource(resName, attribs, data));
                    }
                }
            }
            else
            {
                // Replacing binary resources inside a .resources file
                Debug.Fail("NYI");

                Assembly srcAssembly = Assembly.LoadFrom(src);
                Assembly dstAssembly = Assembly.LoadFrom(dest);
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
