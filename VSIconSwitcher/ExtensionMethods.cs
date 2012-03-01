using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using Mono.Cecil;
using System.Collections;
using System.Diagnostics;
using System.IO;

namespace VSIconSwitcher
{
    public static class ExtensionMethods
    {
        public static bool ContainsResource(this ResourceReader @this, string resName)
        {
            try
            {
                string dummy1;
                byte[] dummy2;
                @this.GetResourceData(resName, out dummy1, out dummy2);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        public static UnmanagedMemoryStream FindResource(this ResourceReader @this, string resName)
        {
            IDictionaryEnumerator e = @this.GetEnumerator();
            while (e.MoveNext())
            {
                string key = e.Key as string;
                Debug.Assert(key != null);

                if (key.Equals(resName))
                {
                    UnmanagedMemoryStream data = e.Value as UnmanagedMemoryStream;
                    Debug.Assert(data != null);
                    return data;
                }
            }
            Debug.Fail("Should never get here. Couldn't find a matching resource for name " + resName);
            return null;
        }

        public static void ReplaceResource(this AssemblyDefinition @this, string resName, byte[] newData)
        {
            Resource resourceToReplace = @this.MainModule.Resources.Single(r => r.Name.Equals(resName, StringComparison.OrdinalIgnoreCase));
            ManifestResourceAttributes attribs = resourceToReplace.Attributes;
            int index = @this.MainModule.Resources.IndexOf(resourceToReplace);

            @this.MainModule.Resources.RemoveAt(index);
            @this.MainModule.Resources.Insert(index, new EmbeddedResource(resName, attribs, newData));
        }
    }
}
