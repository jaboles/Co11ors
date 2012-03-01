using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace VSIconSwitcher
{
    public class ReplacementList
    {
        private static IDictionary<string, IEnumerable<AssetReplacement>> s_lists = new Dictionary<string, IEnumerable<AssetReplacement>>();

        public static IEnumerable<AssetReplacement> VS11Ultimate { get { return GetOrRead("VS11Ultimate.txt"); } }

        public static IEnumerable<AssetReplacement> GetOrRead(string name)
        {
            if (s_lists.ContainsKey(name))
            {
                return s_lists[name];
            }

            List<AssetReplacement> list = new List<AssetReplacement>();

            StreamReader sr = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream(string.Format("VSIconSwitcher.{0}", name)));
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine().Trim();
                if (string.IsNullOrEmpty(line))
                    line = sr.ReadLine().Trim();

                if (line.StartsWith("--") || string.IsNullOrWhiteSpace(line))
                    continue;

                string[] stringParts = line.Split(';');
                string typeIdentifier = stringParts[0];
                string fileName = stringParts[1];
                ResourceType resType = 0;

                if (typeIdentifier.Equals("File") || typeIdentifier.Equals("F"))
                {
                    list.Add(new FileReplacement(fileName));
                }
                else if (typeIdentifier.Equals("ManagedDefault") || typeIdentifier.Equals("G"))
                {
                    string resourceNamePattern = stringParts[2];
                    list.Add(new ManagedResourceReplacement(fileName, resourceNamePattern));
                }
                else if (typeIdentifier.Equals("Managed") || typeIdentifier.Equals("M"))
                {
                    string dstFilename = stringParts[2];
                    string resourcesFilename = stringParts.Length > 3? stringParts[3] : null;
                    string resourceNamePattern = stringParts.Length > 4?  stringParts[4] : null;

                    if (string.IsNullOrWhiteSpace(dstFilename)) dstFilename = null;

                    list.Add(new ManagedResourceReplacement(fileName, dstFilename, resourcesFilename, resourceNamePattern));
                }
                else
                {
                    string idString = stringParts[2];
                    if (typeIdentifier.Equals("Icon") || typeIdentifier.Equals("I"))
                    {
                        resType = ResourceType.Icon;
                    }
                    else if (typeIdentifier.Equals("Bitmap") || typeIdentifier.Equals("B"))
                    {
                        resType = ResourceType.Bitmap;
                    }
                    else
                    {
                        Debug.Fail("Unknown type identifier: " + typeIdentifier);
                    }

                    List<string> ids = new List<string>();
                    foreach (string s in idString.Split(','))
                    {
                        string idStr = s.Trim();
                        if (idStr.Contains('-'))
                        {
                            string[] parts = idStr.Split('-');
                            int lower = Convert.ToInt32(parts[0].Trim());
                            int upper = Convert.ToInt32(parts[1].Trim());
                            Debug.Assert(lower <= upper, "Range should be specified <lower>-<upper>");
                            for (int i = lower; i <= upper; i++)
                            {
                                ids.Add(i.ToString());
                            }
                        }
                        else
                        {
                            ids.Add(idStr);
                        }
                    }

                    list.Add(new NativeResourceReplacement(fileName, resType, ids.ToArray()));
                }
            }

            s_lists[name] = list;
            return list;
        }
    }
}
