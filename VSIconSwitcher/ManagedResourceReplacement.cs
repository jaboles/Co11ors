using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VSIconSwitcher
{
    public class ManagedResourceReplacement : AssetReplacement
    {
        private string m_resourceNamePattern;

        public ManagedResourceReplacement(string filename, string resourceNamePattern)
            : base(filename)
        {
            m_resourceNamePattern = filename;
        }

        public override void DoReplace()
        {
        }
    }
}
