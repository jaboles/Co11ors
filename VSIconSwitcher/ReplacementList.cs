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
        private static IDictionary<string, IEnumerable<ResourceReplacement>> s_lists = new Dictionary<string, IEnumerable<ResourceReplacement>>(); 

        public static IEnumerable<ResourceReplacement> VS11Ultimate { get { return GetOrRead("VS11Ultimate.txt"); } }

        public static IEnumerable<ResourceReplacement> GetOrRead(string name)
        {
            if (s_lists.ContainsKey(name))
            {
                return s_lists[name];
            }

            List<ResourceReplacement> list = new List<ResourceReplacement>();

            StreamReader sr = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream(string.Format("VSIconSwitcher.{0}", name)));
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine().Trim();
                if (string.IsNullOrEmpty(line))
                    line = sr.ReadLine().Trim();

                string[] stringParts = line.Split(';');
                string fileName = stringParts[1];
                string typeIdentifier = stringParts[0];
                string idString = stringParts[2];
                ResourceType resType = 0;

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

                List<int> ids = new List<int>();
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
                            ids.Add(i);
                        }
                    }
                    else
                    {
                        ids.Add(Convert.ToInt32(idStr));
                    }
                }

                list.Add(new ResourceReplacement(fileName, resType, ids.ToArray()));
            }

            s_lists[name] = list;
            return list;
        }
    }
}
