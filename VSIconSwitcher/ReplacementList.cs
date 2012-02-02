using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VSIconSwitcher
{
    public static class ReplacementList
    {
        public static ResourceReplacement[] VS11Ultimate = new ResourceReplacement[] {
            new ResourceReplacement("devenv.exe", ResourceType.Icon, 1200, 1202, 1203, 6826),
            new ResourceReplacement("1033\\cmddefui.dll", ResourceType.Bitmap,
                new Range(2001, 2008),
                new Range(2012, 2017),
                new Range(2020),
                new Range(2022, 2035)),
            new ResourceReplacement("1033\\ExtWizrdUI.dll", ResourceType.Icon, 174),
            new ResourceReplacement("1033\\ExtWizrdUI.dll", ResourceType.Icon, 174),
            new ResourceReplacement("1033\\msenvui.dll", ResourceType.Icon, 1200, 4911),
            new ResourceReplacement("1033\\msenvui.dll", ResourceType.Bitmap, 300, 304, 305, 6650, 6651),
            new ResourceReplacement("msenv.dll", ResourceType.Bitmap, 327, 328, 1268, 1423, 1425, 1427, 6000, 6100, 7003),
            new ResourceReplacement("msenv.dll", ResourceType.Icon,
                new Range(300),
                new Range(1207),
                new Range(6807, 6808),
                new Range(6810, 6827),
                new Range(6830)),
            
        };
    }
}
