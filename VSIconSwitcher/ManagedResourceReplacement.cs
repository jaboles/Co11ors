using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;

namespace VSIconSwitcher
{
    public class ManagedResourceReplacement : AssetReplacement
    {
        private string m_resourceNamePattern;
        private string m_resourcesFilename;

        public ManagedResourceReplacement(string filename, string resourcesFilename, string resourceNamePattern)
            : base(filename)
        {
            m_resourcesFilename = resourcesFilename;
            m_resourceNamePattern = resourceNamePattern;
        }

        public ManagedResourceReplacement(string filename, string resourceNamePattern)
            : base(filename)
        {
            m_resourcesFilename = Path.GetFileNameWithoutExtension(filename) + ".g";
            m_resourceNamePattern = resourceNamePattern;
        }

        public override void CopyResources(string src, string dest)
        {


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
    }
}
